import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { runProcess, dataDir } from './download.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SCRIPT = path.join(__dirname, '..', 'python', 'extract_doc.py');

function findPythonCandidates() {
  const out = [];
  const rtRoot = path.join(os.homedir(), '.dsh', 'dsh-runtimes');
  if (fs.existsSync(rtRoot)) {
    for (const d of fs.readdirSync(rtRoot)) {
      const p = path.join(rtRoot, d, 'dependencies', 'python', 'python.exe');
      if (fs.existsSync(p)) out.push(p);
    }
  }
  out.push('python', 'python3');
  return out;
}

async function resolvePython() {
  const errors = [];
  for (const p of findPythonCandidates()) {
    const r = await runProcess(p, ['--version'], { timeout: 15000 });
    if (r.code === 0) return p;
    errors.push(`${p}: ${r.stderr.trim() || 'exit ' + r.code}`);
  }
  throw new Error('未找到可用的 Python 运行时：\n' + errors.join('\n'));
}

async function ensurePymupdf(python, target) {
  const check = await runProcess(python, ['-c', 'import fitz'], { timeout: 30000 });
  if (check.code === 0) return;
  fs.mkdirSync(target, { recursive: true });
  const pip = await runProcess(python, ['-m', 'pip', 'install', '--disable-pip-version-check', '--target', target, 'pymupdf'], {
    timeout: 600000,
  });
  if (pip.code !== 0) {
    throw new Error('PyMuPDF 自动安装失败：' + (pip.stderr || pip.stdout).trim().slice(-2000));
  }
}

/**
 * 提取文档文本与公式。
 * @param {{path:string, include_notes?:boolean}} args
 * @param {{signal?:AbortSignal}} exec
 */
export async function extractDocument(args, exec = {}) {
  const filePath = args.path;
  if (!filePath || typeof filePath !== 'string') {
    return { ok: false, error: '参数 path（文档绝对路径）不能为空' };
  }
  const abs = path.resolve(filePath);
  if (!fs.existsSync(abs)) {
    return { ok: false, error: `文件不存在：${abs}` };
  }

  let python;
  try {
    python = await resolvePython();
  } catch (e) {
    return { ok: false, error: e.message };
  }

  const pyDir = path.join(dataDir(), 'py');
  const ext = path.extname(abs).toLowerCase();
  if (ext === '.pdf') {
    try {
      await ensurePymupdf(python, pyDir);
    } catch (e) {
      return { ok: false, error: e.message };
    }
  }

  const env = {
    ...process.env,
    PYTHONUTF8: '1',
    PYTHONPATH: [pyDir, process.env.PYTHONPATH].filter(Boolean).join(path.delimiter),
  };
  const flag = args.include_notes === false ? 'no-notes' : 'with-notes';
  const r = await runProcess(python, [SCRIPT, abs, flag], { env, signal: exec.signal, timeout: 300000 });

  if (r.code !== 0) {
    return { ok: false, error: '提取失败', detail: (r.stderr || r.stdout).trim().slice(-4000) };
  }
  try {
    return JSON.parse(r.stdout);
  } catch {
    return { ok: false, error: '解析提取结果失败', detail: r.stdout.trim().slice(-4000) };
  }
}
