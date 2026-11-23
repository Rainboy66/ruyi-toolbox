# 架构与产品决策

## 模块

`app.py` 负责入口、任务队列和功能路由；`desktop_ui.py`、`home_surface.py`、`window_paint.py` 负责原生 Tk 界面和 Windows 绘制；`desktop_integration.py` 处理 DPI、设置及显式桌面入口。

`office_engine.py` 选择本机办公引擎，`native_office.py` 与相应 PowerShell 脚本使用任务自有会话导出。`external_resources.py` 在临时副本中处理资源路径、缺失资源和缓存，不修改用户原件。

`pdf_engine.py` 路由三种反向输出；`pdf_docx_*` 分析段落、目录、表格、公式及位置，`pdf_xlsx_layout.py`、`pdf_pptx_*` 分别重建表格和幻灯片对象；`pdf_word_images.py` 只处理图片编码，不修改显示几何。

`image_engine.py` 处理图片格式、尺寸和体积；`id_photo_engine.py` 使用 MODNet 透明度、边缘细化及旧底色校正，本地合成新背景。预设颜色集中在 `id_photo_backgrounds.py`。

`pdf_organizer/` 使用统一的 PageRef 和 OutputPlan；不可变方案、来源校验、50 步撤销重做及 JSON 序列化由 model.py 负责。exporter/composition/overlays 根据同一方案生成 PDF，预设不另建第二套页面模型。缩略图和预览使用后台进程、可见区域与有限缓存，避免将所有页面同步渲染到 Tk 主线程。

## 取舍

- PDF 页面重放与可编辑语义重建是两件事：正文/表格优先原生编辑，缺失数学字形等复杂公式允许局部图片保真，不能将整页截图说成完全可编辑。
- 图片体积优化以透明度、尺寸、颜色分布及源流可复用性分支处理，不通过改变图片像素尺寸掩盖编码问题。
- PDF 注释和表单不能仅因视觉可见就判定保留；会破坏交互语义的拼版明确拒绝，批量继续其他文件。
- 源码测试、实际 Office/WPS 渲染、最终 EXE 与人工交互分别验收，不互相替代。

## 人与 AI 的分工

项目发起人提出实际场景、需求优先级和验收标准，选择视觉与交互方向，使用真实文件发现并反馈问题。Codex、DeepSeek Harness 等参与辅助编码、方案讨论、排查与自动化验证。不把 AI 辅助实现描述为个人独立编写全部代码，也不把开源模型描述为自研模型。软件运行不依赖大模型服务。
