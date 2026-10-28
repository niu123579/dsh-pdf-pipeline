import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { spawn, spawnSync } from 'node:child_process';

export const PANDOC_VERSION = '3.12.1';
export const TECTONIC_VERSION = '0.17.0';

// 下载地址按优先级排列：直连 + 两个可用的 GitHub 镜像。
const PANDOC_ASSET = `pandoc-${PANDOC_VERSION}-windows-x86_64.zip`;
const TECTONIC_ASSET = `tectonic-${TECTONIC_VERSION}-x86_64-pc-windows-msvc.zip`;

// 镜像优先：gh-proxy.com 实测最快，GitHub 直连在部分网络下大文件会 reset，故放最后兜底。
export const PANDOC_URLS = [
  `https://gh-proxy.com/https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/${PANDOC_ASSET}`,
  `https://ghfast.top/https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/${PANDOC_ASSET}`,
  `https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/${PANDOC_ASSET}`,
];

export const TECTONIC_URLS = [
  `https://gh-proxy.com/https://github.com/tectonic-typesetting/tectonic/releases/download/tectonic%40${TECTONIC_VERSION}/${TECTONIC_ASSET}`,
  `https://ghfast.top/https://github.com/tectonic-typesetting/tectonic/releases/download/tectonic%40${TECTONIC_VERSION}/${TECTONIC_ASSET}`,
  `https://github.com/tectonic-typesetting/tectonic/releases/download/tectonic%40${TECTONIC_VERSION}/${TECTONIC_ASSET}`,
];

/** 插件数据目录：二进制、缓存、导出产物。 */
export function dataDir() {
  return path.join(os.homedir(), '.dsh', 'pdf-pipeline');
}
export function binDir() {
  return path.join(dataDir(), 'bin');
}
export function exportsDir() {
  return path.join(dataDir(), 'exports');
}

/** 带重试的 fetch 下载，逐 URL 尝试，成功才返回。 */
async function downloadFile(urls, dest) {
  const errors = [];
  for (const url of urls) {
    try {
      await fetchToFile(url, dest);
      return;
    } catch (err) {
      errors.push(`${url} -> ${err.message}`);
      try { fs.rmSync(dest, { force: true }); } catch { /* ignore */ }
    }
  }
  throw new Error(`下载失败：\n${errors.join('\n')}`);
}

async function fetchToFile(url, dest) {
  const res = await fetch(url, { redirect: 'follow', signal: AbortSignal.timeout(300000) });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const total = Number(res.headers.get('content-length') || 0);
  const tmp = `${dest}.part`;
  const ws = fs.createWriteStream(tmp);
  let received = 0;
  try {
    for await (const chunk of res.body) {
      received += chunk.length;
      if (!ws.write(Buffer.from(chunk))) {
        await new Promise((r) => ws.once('drain', r));
      }
    }
  } finally {
    await new Promise((resolve, reject) => ws.end((e) => (e ? reject(e) : resolve())));
  }
  if (total > 0 && received < total) {
    fs.rmSync(tmp, { force: true });
    throw new Error(`下载不完整：${received}/${total} 字节`);
  }
  fs.renameSync(tmp, dest);
}

/** 用 PowerShell Expand-Archive 解压 zip（Windows 原生，可靠）。 */
function extractZip(zipPath, destDir) {
  fs.mkdirSync(destDir, { recursive: true });
  const esc = (s) => s.replace(/'/g, "''");
  const cmd = `Expand-Archive -LiteralPath '${esc(zipPath)}' -DestinationPath '${esc(destDir)}' -Force`;
  const r = spawnSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', cmd], {
    encoding: 'utf8', timeout: 300000, windowsHide: true,
  });
  if (r.status !== 0) {
    throw new Error(`解压失败（exit ${r.status}）：${r.stderr || r.stdout || '未知错误'}`);
  }
}

function findFile(dir, name) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      const found = findFile(p, name);
      if (found) return found;
    } else if (entry.name.toLowerCase() === name.toLowerCase()) {
      return p;
    }
  }
  return null;
}

/**
 * 确保某个可执行文件存在：不存在则下载、解压、定位并复制到 bin 目录，缓存复用。
 */
export async function ensureBinary(name, urls, exeName) {
  const bdir = binDir();
  fs.mkdirSync(bdir, { recursive: true });
  const target = path.join(bdir, exeName);
  if (fs.existsSync(target) && fs.statSync(target).size > 0) return target;

  const tmp = path.join(os.tmpdir(), `dsh-pdf-${name}-${Date.now()}`);
  fs.mkdirSync(tmp, { recursive: true });
  const zip = path.join(tmp, `${name}.zip`);
  await downloadFile(urls, zip);
  const exdir = path.join(tmp, 'extracted');
  extractZip(zip, exdir);
  const found = findFile(exdir, exeName);
  if (!found) throw new Error(`压缩包中未找到 ${exeName}`);
  fs.copyFileSync(found, target);
  fs.rmSync(tmp, { recursive: true, force: true });
  return target;
}

/**
 * 运行外部进程，返回 { code, stdout, stderr }；支持 AbortSignal 取消与超时。
 */
export function runProcess(cmd, args, { cwd, env, signal, timeout = 600000 } = {}) {
  return new Promise((resolve) => {
    const child = spawn(cmd, args, {
      cwd, env: env ?? process.env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (d) => { stdout += d.toString(); });
    child.stderr.on('data', (d) => { stderr += d.toString(); });

    const kill = () => { try { child.kill(); } catch { /* ignore */ } };
    const onAbort = () => kill();
    if (signal) signal.addEventListener('abort', onAbort, { once: true });

    const timer = setTimeout(kill, timeout);
    let settled = false;
    const finish = (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (signal) signal.removeEventListener('abort', onAbort);
      resolve({ code, stdout, stderr });
    };
    child.on('error', (err) => {
      stderr += `\n进程启动失败: ${err.message}`;
      finish(-1);
    });
    child.on('close', (code) => finish(code ?? -1));
  });
}
