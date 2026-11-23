# 人像模型

- 原项目：https://github.com/ZHKKKe/MODNet
- ONNX 发布者：https://huggingface.co/Xenova/modnet
- 固定版本：e231692d2977d62c7be22830f92d9f8a2f0737f7
- 下载：https://huggingface.co/Xenova/modnet/resolve/e231692d2977d62c7be22830f92d9f8a2f0737f7/onnx/model.onnx
- 文件大小：25,888,640 字节
- SHA256：07c308cf0fc7e6e8b2065a12ed7fc07e1de8febb7dc7839d7b7f15dd66584df9
- 许可：Apache-2.0；原文在 third_party/licenses/MODNet/LICENSE.txt。

本仓库不含权重。源码首次处理照片时，程序在 `%LOCALAPPDATA%\Ruyi\models` 获取并校验 `modnet.onnx`，随后本地推理。网络请求只获取模型，不发送照片。该模型并非本项目自行训练。

离线配置：从固定下载地址获取文件，核对大小和 SHA256 后重命名为 `modnet.onnx`，放入自选目录。在启动应用的同一个 PowerShell 设置：

```powershell
$env:RUYI_MODEL_DIR = 'D:\RuyiModels'
.\.venv\Scripts\python.exe app.py
```

环境变量指向包含文件的目录。打包脚本会将校验通过的同一模型嵌入 EXE；新编译的包仍须单独检查、收集许可并验收。
