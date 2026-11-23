param([string]$Python = 'python')
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Push-Location -LiteralPath $repo
try {
    $model = (& $Python -B -c 'from id_photo_engine import _model_path; print(_model_path())')
    if ($LASTEXITCODE -ne 0) { throw '模型获取或校验失败。' }
    $model = ([string]$model).Trim()
    $modelFolder = Split-Path -Parent $model
    & $Python -m PyInstaller --noconfirm --clean --onefile --windowed `
        --name '如意' --icon assets/ruyi.ico --version-file version_info.txt `
        --add-data 'assets;assets' `
        --add-data 'native_export.ps1;.' --add-data 'native_word_export.ps1;.' `
        --add-data 'native_wps_export.ps1;.' --add-data 'native_excel_export.ps1;.' `
        --add-data ($modelFolder + ';models') `
        --collect-binaries onnxruntime --collect-data reportlab `
        --copy-metadata pypdf --copy-metadata reportlab --exclude-module pandas `
        --distpath dist --workpath .build/work --specpath .build app.py
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller 构建失败。' }
    Copy-Item -LiteralPath LICENSE -Destination dist/LICENSE
    Copy-Item -LiteralPath THIRD_PARTY_NOTICES.md -Destination dist/THIRD_PARTY_NOTICES.md
    Copy-Item -LiteralPath third_party/licenses -Destination dist/licenses -Recurse -Force
    Write-Output '输出位于 dist。分发前需核对实际依赖许可并完成新 EXE 验收。'
} finally {
    Pop-Location
}
