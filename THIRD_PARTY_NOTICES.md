# 第三方组件与许可

根目录 MIT 只覆盖本项目有权许可的自有代码，不替换依赖、模型、字体或原作者的许可。

| 组件 | 用途 | 许可说明 |
| --- | --- | --- |
| Python / Tcl-Tk | 运行时与桌面窗口 | Python PSF 及 Tcl/Tk 原条款 |
| Pillow | 图片处理 | HPND 等原条款 |
| pdfplumber / pdfminer.six | PDF 结构与文字 | MIT |
| pypdfium2 / PDFium | PDF 渲染 | pypdfium2 为 Apache-2.0 或 BSD-3-Clause；PDFium 及内含组件见完整声明 |
| pypdf / ReportLab | PDF 写入、拼版与模板 | BSD 系许可，见各组件原文 |
| python-docx / openpyxl / python-pptx | Office 文件写出 | MIT |
| NumPy / ONNX Runtime | 本地人像推理 | NumPy 为 BSD 系，ONNX Runtime 为 MIT，并含各自第三方声明 |
| MODNet / Xenova ONNX 转换 | 人像模型 | Apache-2.0；来源与权重校验见 docs/MODEL.md |
| Lucide | 功能图标 | ISC，保留 assets/Lucide-LICENSE.txt |
| olefile / lxml 及传递依赖 | 文档解析 | 以原包声明为准 |
| PyInstaller | 可选打包工具 | GPL 及其分发例外；不将其误写为 MIT |

`third_party/licenses/` 保留既有 Windows 1.1.0 发布包收集的许可原文，包括 PDFium 及 ONNX Runtime 的内含组件说明；未把组件代码或模型权重复制进源码仓库。版本变化后应重新核查实际分发物，不能用旧声明代替新组件许可。

第三方办公软件及系统字体没有捆绑进本源码。Microsoft Office、WPS、LibreOffice 需要使用者另行安装并按各自条款使用。B08 品牌图形为项目 AI 辅助设计素材，不属于 Lucide 图标库。
