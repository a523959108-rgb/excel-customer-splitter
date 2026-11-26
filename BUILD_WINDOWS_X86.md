# 在 Intel/AMD Windows 上构建

1. 在目标电脑安装 Windows x64 Python 3.11 或更高版本，安装时勾选 **Add Python to PATH**。必须使用完整安装版 Python，不能使用 embeddable zip。
2. 将本目录复制到目标电脑，例如 `C:\Excel客户拆分`。
3. 打开 PowerShell，进入该目录并运行：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\build_windows.ps1
```

构建脚本会检查 Tkinter 和 Python 架构，安装依赖并生成：

```text
dist\Excel客户拆分\Excel客户拆分.exe
```

发布时复制整个 `dist\Excel客户拆分` 文件夹，包括 `_internal`、`regions.json` 和 `customer_config.json`。脚本使用 x64 Python 构建并运行桌面自检。Windows 11 ARM 虚拟机也可以通过 x64 兼容层运行该程序。

要构建中文安装向导，在完成上述构建后使用 Inno Setup 6.6 或更高的 6.x 版本运行 `.\build_installer.ps1`。产物位于 `dist\installer\ExcelCustomerSplitter-Setup-x64.exe`。GitHub Actions 的 **Build Windows x64** 工作流会额外验证安装、覆盖安装以及卸载后的数据库内容。
