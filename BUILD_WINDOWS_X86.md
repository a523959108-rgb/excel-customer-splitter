# 在 Intel/AMD Windows 上构建

1. 在目标电脑安装 Windows Python 3.11 或更高版本，安装时勾选 **Add Python to PATH**。必须使用完整安装版 Python，不能使用 embeddable zip。
2. 将本目录复制到目标电脑，例如 `C:\Excel客户拆分`。
3. 打开 PowerShell，进入该目录并运行：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\build_windows.ps1
```

构建脚本会检查 Tkinter 和 Python 架构，安装依赖并生成：

```text
dist\Excel客户拆分.exe
```

请把 `regions.json` 和 `customer_config.json` 放在 exe 同目录。当前 Parallels 中的 Windows 是 ARM64，不能用它生成 Intel/AMD exe；必须在目标 Intel/AMD Windows 上执行构建。
