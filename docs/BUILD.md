# Windows 打包

推荐 Windows x64 / Python 3.12。在仓库根目录：

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/build_windows.ps1 -Python .\.venv\Scripts\python.exe
```

脚本读取当前版本，获取并校验模型，使用 PyInstaller 单文件 GUI 模式及现有 ICO，复制 MIT 和已有组件许可。输出在 `dist/`，中间文件在 `.build/`，两者均被 Git 忽略。

本次公开整理没有重新构建 EXE；提供的构建脚本只完成参数与语法检查。新环境构建后需要重新执行实际转换、模型启动、图标、Office 引擎和界面验收，不把历史二进制的验证当作新构建已通过。若依赖版本改变，必须重新收集实际安装包的完整许可，不能直接使用仓库内旧版声明作为最终分发依据。

既有 1.1.0 下载包由已验收 EXE 重新整理为 ZIP，EXE 字节不变。程序尚未代码签名，哈希校验与签名是不同机制。
