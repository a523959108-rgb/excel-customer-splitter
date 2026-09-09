$ErrorActionPreference = "Stop"

# Always resolve files relative to this script. This lets the source package
# build correctly when launched from PowerShell, Explorer, or another folder.
Set-Location -LiteralPath $PSScriptRoot

$python = (Get-Command python -ErrorAction Stop).Source
$pythonArch = & $python -c "import platform; print(platform.machine())"
Write-Host "Python 架构：$pythonArch"
if ($pythonArch -eq "ARM64") {
    throw "当前 Python 是 ARM64。请在 Intel/AMD Windows 上安装对应架构的 Python 后再构建；PyInstaller 不能跨架构生成 x64 exe。"
}

Write-Host "检查 Python GUI 组件..."
& $python -c "import tkinter; print('Tk', tkinter.TkVersion)"
if ($LASTEXITCODE -ne 0) {
    throw "当前 Python 没有 tkinter/Tcl-Tk。请安装完整的 Windows Python（不要使用 embeddable zip）后再运行此脚本。"
}
& $python -m pip install --upgrade pip
& $python -m pip install -r requirements.txt
& $python -m PyInstaller --noconfirm --clean --onefile --windowed --name Excel客户拆分 app.py
Copy-Item regions.json dist/regions.json -Force
if (Test-Path customer_config.json) { Copy-Item customer_config.json dist/customer_config.json -Force }
Write-Host "构建完成：dist/Excel客户拆分.exe（架构：$pythonArch）"
