# 客户 Excel 拆分工具（MVP）

这是第一版 Windows 桌面程序：将多个 `.xlsx` 导入 SQLite 暂存库，汇总字段后由用户选择客户字段、分区字段和导出字段；程序按表头自动识别年 / 月 / 日，并在日历中选择导出日期范围，再按以下结构生成文件：

```text
输出目录/日期/分区/客户.xlsx
```

## Windows 构建

在目标 Windows 机器上安装带 Tcl/Tk 的完整 Python 3.11 或更高版本，然后在项目目录执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\build_windows.ps1
```

最终文件位于 `dist/Excel客户拆分.exe`。`regions.json` 需要和 exe 放在同一目录，用户可以复制并编辑分区配置。

不要使用 Python embeddable zip 打包：它通常不包含 `tkinter`，构建脚本会在开始时检查并拒绝这种环境。PyInstaller 不能跨架构编译，程序必须在目标架构的 Windows 和 Python 中构建。Intel/AMD 64 位电脑请使用 AMD64/x64 Python；如果目标真的是 32 位 Windows，请使用 32 位 x86 Python。当前 Mac 上的 Parallels Windows 是 ARM64，只能用于 ARM64 验证，不能生成 Intel/AMD exe。

## GitHub Actions 在线构建

仓库包含 `.github/workflows/build-windows.yml`。推送到 `main` 分支，或在 GitHub 的 **Actions → Build Windows x64 → Run workflow** 手动运行后，打开对应运行记录，在 **Artifacts** 下载 `ExcelCustomerSplitter-windows-x64.zip`。压缩包内包含 x64 的 exe、`regions.json`、`customer_config.json` 和 README，可直接放到 Windows x64 电脑使用。

## 使用流程

1. 选择多个 Excel 文件。
2. 点击“扫描表头并导入 SQLite”。程序默认读取每个文件的第一个工作表，并自动寻找前 30 行中的第一个有效表头行。
3. 选择导出字段、客户字段和分区字段。程序按表头自动识别“年”“月”“日”列。
4. 可在 `customer_config.json` 的 `customers` 数组中预设客户名称；加载后勾选要导出的客户。名单为空或配置文件不存在时导出全部客户。
5. 导入完成后，日历只启用数据日期范围内的日期；第一次点击选择开始日期，第二次点击选择结束日期，按闭区间筛选。选中一段日期后，同一客户在这段时间内的数据会写入同一个客户文件，并放在一个日期范围文件夹中；点击清除日期范围则导出全部日期。
6. 选择分区配置和输出目录。
7. 点击“开始拆分导出”。

程序按批次写入 SQLite，导出时使用磁盘临时分组表和 `openpyxl` 流式工作簿，避免把全量数据放进 Python 内存，适合千万级总行数（实际耗时和磁盘空间取决于列数、客户数和输出文件数量）。单个客户超过 Excel 单表行数限制时会自动生成 `_part2.xlsx` 等分卷文件。导出过程中可以点击“结束任务并释放资源”，程序会协作停止当前任务、清理临时表和工作簿，并重新建立 SQLite 连接。
