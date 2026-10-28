import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import {
  PANDOC_URLS, TECTONIC_URLS, ensureBinary, runProcess, exportsDir,
} from './download.js';

const HEADER_TEX = String.raw`
\usepackage{amsmath,amssymb}
\usepackage[version=4]{mhchem}
\usepackage{chemfig}
\usepackage{xeCJK}
\setCJKmainfont{Microsoft YaHei}
\setCJKsansfont{Microsoft YaHei}
\setCJKmonofont{Microsoft YaHei}
`;

function sanitizeFilename(s) {
  return (s || 'document').replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim().slice(0, 120) || 'document';
}

/**
 * 导出 Markdown → PDF。
 * @param {{content:string, output_path?:string, title?:string}} args
 * @param {{signal?:AbortSignal}} exec
 */
export async function exportPdf(args, exec = {}) {
  const content = args.content;
  if (!content || typeof content !== 'string' || !content.trim()) {
    return { ok: false, error: '参数 content（Markdown 正文）不能为空' };
  }
  const title = sanitizeFilename(args.title || 'document');
  const outputPath = args.output_path
    ? path.resolve(args.output_path)
    : path.join(exportsDir(), `${Date.now()}-${title}.pdf`);

  let pandoc; let tectonic;
  try {
    pandoc = await ensureBinary('pandoc', PANDOC_URLS, 'pandoc.exe');
    tectonic = await ensureBinary('tectonic', TECTONIC_URLS, 'tectonic.exe');
  } catch (err) {
    return { ok: false, error: `依赖下载失败（首次使用需联网下载 Pandoc/Tectonic）：${err.message}` };
  }

  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-pdf-'));
  try {
    const mdPath = path.join(work, 'input.md');
    const headerPath = path.join(work, 'header.tex');
    const texPath = path.join(work, 'output.tex');
    const builtPdf = path.join(work, 'output.pdf');
    fs.writeFileSync(mdPath, content, 'utf8');
    fs.writeFileSync(headerPath, HEADER_TEX, 'utf8');

    // 第 1 步：Pandoc 把 Markdown 转成独立 LaTeX。
    const mdRes = await runProcess(pandoc, [
      mdPath, '--standalone', '--include-in-header', headerPath,
      '-o', texPath,
      '-V', 'geometry=margin=2.5cm',
      '-V', `title=${title}`,
      '-V', 'colorlinks=true',
    ], { cwd: work, signal: exec.signal, timeout: 120000 });

    if (mdRes.code !== 0 || !fs.existsSync(texPath)) {
      return {
        ok: false, error: 'Pandoc 转换 LaTeX 失败',
        detail: (mdRes.stderr || mdRes.stdout).trim().slice(-4000),
      };
    }

    // 第 2 步：Tectonic 编译 LaTeX → PDF（首次会下载宏包缓存）。
    const texRes = await runProcess(tectonic, ['output.tex'], {
      cwd: work, signal: exec.signal, timeout: 600000,
    });

    if (texRes.code !== 0 || !fs.existsSync(builtPdf)) {
      return {
        ok: false, error: 'Tectonic 编译 PDF 失败',
        detail: (texRes.stderr || texRes.stdout).trim().slice(-4000),
      };
    }

    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.copyFileSync(builtPdf, outputPath);
    return {
      ok: true,
      pdfPath: outputPath,
      pdfSize: fs.statSync(outputPath).size,
      title,
      note: 'PDF 已生成。如需让用户直接查看，可用 present 工具呈现该文件。',
    };
  } finally {
    try { fs.rmSync(work, { recursive: true, force: true }); } catch { /* ignore */ }
  }
}
