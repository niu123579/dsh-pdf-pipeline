# dsh-pdf-pipeline

面向 DeepSeek Harness（DSH）的 PDF 输出优化插件：把汇总内容渲染成排版精美、公式正确的高质量 PDF，并从含公式的 PPT / PDF / Word 中提取文本与公式。

> 大学生 / 科研场景神器：LaTeX 数学、mhchem 化学反应方程式、chemfig 有机结构式，中文正文，全部一条命令搞定。

## 功能

提供两个 Agent 工具：

### `export_pdf` — Markdown → 精美 PDF

基于 **Pandoc + Tectonic（XeTeX）** 引擎，中文为主、公式正确渲染：

- `$...$` 行内数学、`$$...$$` 块级数学
- `\ce{...}` 化学反应方程式（mhchem）
- `\chemfig{...}` 有机结构式（chemfig）
- 表格、列表、加粗 / 斜体 / 代码块
- 中文使用微软雅黑（`xeCJK`）

首次调用自动联网下载 Pandoc 与 Tectonic（约 60MB，免管理员），之后缓存复用。生成后返回 PDF 绝对路径，可配合 `present` 工具直接呈现给用户。

### `extract_document` — 从文档提取文本与公式

- 支持 `.pptx`（正文 + 演讲者备注 + 文本框 / 表格 / 组合内的公式）、`.docx`、`.pdf`
- Office 的 OMML 原生公式自动转换为 LaTeX
- PDF 使用 PyMuPDF 提取文本（首次按需安装）；图片式公式无法 OCR，会启发式标注疑似公式行
- 返回结构化 JSON：每页 / 每张幻灯片的文本与识别出的公式 LaTeX

## 安装

```bash
# 1. 克隆仓库
git clone https://github.com/niu123579/dsh-pdf-pipeline.git

# 2. 在 DSH 中安装本地目录 bundle
#    （对话中让 Agent 执行 plugin_manager install_bundle，target 指向克隆目录）
```

也可直接让 DSH 里的 Agent 调用 `plugin_manager` 的 `install_bundle`，target 填本仓库的绝对路径。

## 使用示例

### 导出 PDF

```
把下面的内容导出成 PDF：
# 高斯积分
$$\int_{-\infty}^{+\infty} e^{-x^2}\,dx = \sqrt{\pi}$$
# 电解水
$$\ce{2H2O -> 2H2 ^ + O2 ^}$$
# 乙醇结构式
$$\chemfig{H-C(-[2]H)(-[6]H)-C(-[2]H)(-[6]H)-O-H}$$
```

### 提取文档

```
解析这个 PPT 的正文、备注和公式：D:\course\lecture3.pptx
```

## 首次运行说明

- `export_pdf` 首次会联网下载 Pandoc（3.12.1）与 Tectonic（0.17.0），缓存到 `~/.dsh/pdf-pipeline/bin`；下载镜像按 `gh-proxy.com → ghfast.top → github.com` 优先级回退。
- `extract_document` 处理 PDF 时首次会 `pip install pymupdf` 到 `~/.dsh/pdf-pipeline/py`。
- Tectonic 首次编译会下载 LaTeX 宏包 bundle（约 47MB，缓存到 `%LOCALAPPDATA%\TectonicProject`）。

## 已知限制

- mhchem 的 `\ce{}` **内部不能放中文**（箭头标签如 `->[通电]` 会报错），中文请写在公式外面：
  `水的电解（通电）：$$\ce{2H2O -> 2H2 ^ + O2 ^}$$`
- PDF 中以图片形式渲染的公式无法 OCR 提取，只能启发式标注疑似公式行。

## 目录结构

```
dsh-pdf-pipeline/
├── index.js                 # 插件入口：注册 export_pdf / extract_document
├── lib/
│   ├── download.js          # 二进制懒加载下载 / 解压 / 进程执行
│   ├── export.js            # Pandoc → LaTeX → Tectonic → PDF
│   └── extract.js           # 调用 Python 提取器（含 PyMuPDF 按需安装）
├── python/
│   └── extract_doc.py       # PPTX / DOCX / PDF 提取 + OMML→LaTeX
├── cordis.patch.yml         # Cordis bundle patch
├── locale/                  # 中英文元数据
├── icon.svg
└── package.json
```

## License

[MIT](./LICENSE)
