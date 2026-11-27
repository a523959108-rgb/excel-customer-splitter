$ErrorActionPreference = "Stop"

# Always resolve files relative to this script. This lets the source package
# build correctly when launched from PowerShell, Explorer, or another folder.
Set-Location -LiteralPath $PSScriptRoot

$python = (Get-Command python -ErrorAction Stop).Source
$pythonArch = & $python -c "import sysconfig; print(sysconfig.get_platform())"
Write-Host "Python 架构：$pythonArch"
if ($pythonArch -ne "win-amd64") {
    throw "当前 Python 架构不适用于 x64 构建。请在 Intel/AMD Windows 上安装对应架构的 Python 后再构建；PyInstaller 不能跨架构生成 x64 exe。"
}

Write-Host "检查 Python GUI 组件..."
& $python -c "import tkinter; print('Tk', tkinter.TkVersion)"
if ($LASTEXITCODE -ne 0) {
    throw "当前 Python 没有 tkinter/Tcl-Tk。请安装完整的 Windows Python（不要使用 embeddable zip）后再运行此脚本。"
}
& $python -m pip install --upgrade pip
& $python -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw "依赖安装失败。" }
& $python -m unittest -q
if ($LASTEXITCODE -ne 0) { throw "完整性测试失败。" }
& $python -m PyInstaller --noconfirm --clean --onedir --windowed --hidden-import _duckdb --name Excel客户拆分 app.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller 构建失败。" }
$smoke = Start-Process -FilePath (Resolve-Path 'dist/Excel客户拆分/Excel客户拆分.exe') -ArgumentList '--self-test' -PassThru
if (-not $smoke.WaitForExit(240000)) { $smoke.Kill(); throw "EXE 自检超时。" }
if ($smoke.ExitCode -ne 0) { throw "EXE 导入、重启和导出自检失败。" }
Copy-Item regions.json dist/Excel客户拆分/regions.json -Force
if (Test-Path customer_config.json) { Copy-Item customer_config.json dist/Excel客户拆分/customer_config.json -Force }
Write-Host "构建完成：dist/Excel客户拆分/Excel客户拆分.exe（架构：$pythonArch）"
