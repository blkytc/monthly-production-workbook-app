# 月度生产表生成器

从已有月度报表生成任意日期范围的新工作簿。程序支持四种固定模板，可修改标题、日期、供应商、行高和列宽，并保留模板样式与公式逻辑。

支持“生产、入库、发运表”“开阳磷矿供贵阳化肥矿石用量”“大水矿石产消存计划一览表”和“开阳大水工业园区调度生产信息”。后三种表支持供应商改名和新增；产消存表可选择新增到“生产日累”或“当日配送量”。旧版 `.xls` 输入会自动调用电脑上已安装的 LibreOffice；Windows 还会自动尝试 Microsoft Excel 或 WPS，输出统一为 `.xlsx`。

## 直接运行

```bash
PYTHONPATH=src python3 -m monthly_workbook
```

运行只需要带 Tkinter 的 Python 3.11 或更高版本。正式发布包中的 `.app` 和 `.exe` 不需要用户安装 Python 或 Codex。普通用户请从 GitHub Release 下载对应系统的 ZIP 包。

## 测试

```bash
PYTHONPATH=src python3 -m unittest discover -s tests -v
```

## macOS 打包

```bash
python3 -m pip install -r requirements-build.txt
sh scripts/build_macos.sh
```

产物位于 `dist/月度生产表生成器.app`。

## Windows 打包

在 Windows 命令提示符中运行：

```bat
py -m pip install -r requirements-build.txt
scripts\build_windows.bat
```

产物位于 `dist\MonthlyWorkbookGenerator.exe`。PyInstaller 不支持从 macOS 交叉生成 Windows 可执行文件，但项目的 GitHub Actions 会自动在 Windows 云端构建，并生成 `MonthlyWorkbookGenerator-Windows.zip`。将代码推送并创建 `v*` 标签后，Windows ZIP 会自动添加到 GitHub Release，用户解压后直接双击 `.exe`。

应用中的“帮助”按钮会打开随程序打包的离线电子帮助文档。
