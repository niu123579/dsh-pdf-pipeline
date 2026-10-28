import { defineTool } from '@deepseek-ai/dsh-tools';
import { exportPdf } from './lib/export.js';
import { extractDocument } from './lib/extract.js';

export const inject = ['tools'];

export function apply(ctx) {
  ctx.tools.register(defineTool({
    name: 'export_pdf',
    description:
      '将 Markdown 内容渲染为排版精美的 PDF（基于 Pandoc + LaTeX/Tectonic 引擎，中文为主 + 公式正确渲染）。'
      + '支持 `$...$` 行内数学、`$$...$$` 块级数学、`\\ce{...}`（mhchem 化学反应方程式，如 \\ce{2H2 + O2 -> 2H2O}）'
      + '与 `\\chemfig{...}`（有机结构式）。首次调用会自动联网下载 Pandoc 与 Tectonic（免管理员、缓存复用）。'
      + '返回生成的 PDF 绝对路径；成功后可用 present 工具把该 PDF 呈现给用户下载查看。',
    parameters: {
      content: {
        type: 'string',
        required: true,
        description: '要渲染的 Markdown 正文（UTF-8）。数学用 $...$ / $$...$$，化学方程式用 \\ce{}，有机结构式用 \\chemfig{}。',
      },
      output_path: {
        type: 'string',
        description: '输出 PDF 的绝对路径；省略则写入插件导出目录并返回该路径。',
      },
      title: {
        type: 'string',
        description: '文档标题（用于 PDF 标题与默认文件名）；省略用 document。',
      },
    },
    output: {
      schema: { type: 'string' },
      render: (_args, value) => [{ type: 'text', text: value }],
    },
    timeoutMs: 1800000,
    async execute(args, exec) {
      return JSON.stringify(await exportPdf(args, exec));
    },
  }));

  ctx.tools.register(defineTool({
    name: 'extract_document',
    description:
      '从文档中提取正文文本、备注，并识别数学/化学公式（Office 的 OMML 公式会尽量转换为 LaTeX）。'
      + '支持 .pptx（含演讲者备注、文本框/表格/组合内公式）、.docx、.pdf。'
      + 'PPTX/DOCX 使用内置 Python + python-pptx/python-docx（无需额外安装）；'
      + 'PDF 文本提取使用 PyMuPDF，首次调用会自动 pip 安装到插件数据目录（需联网）。'
      + 'PDF 中以图片形式渲染的公式无法 OCR 提取，只会用启发式标注疑似公式行。'
      + '返回结构化 JSON：每页/每张幻灯片的文本与识别出的公式 LaTeX。',
    parameters: {
      path: {
        type: 'string',
        required: true,
        description: '要提取的文档绝对路径（.pptx / .docx / .pdf）。',
      },
      include_notes: {
        type: 'boolean',
        description: '是否包含 PPTX 演讲者备注；默认 true。',
      },
    },
    output: {
      schema: { type: 'string' },
      render: (_args, value) => [{ type: 'text', text: value }],
    },
    timeoutMs: 300000,
    async execute(args, exec) {
      return JSON.stringify(await extractDocument(args, exec));
    },
  }));
}
