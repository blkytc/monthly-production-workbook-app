# 月度生产表生成器 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 创建可在 macOS 和 Windows 独立运行的月度生产表桌面生成器。

**Architecture:** `workbook.py` 负责纯工作簿转换，`sizing.py` 负责范围解析和尺寸换算，`app.py` 只负责 Tkinter 交互。PyInstaller 从同一入口分别构建 `.app` 和 `.exe`，离线帮助作为打包资源随应用发布。

**Tech Stack:** Python 3.11+ 标准库、Tkinter、PyInstaller（仅构建时）、unittest

---

### Task 1: 日期、范围与尺寸规则

**Files:**
- Create: `src/monthly_workbook/sizing.py`
- Create: `tests/test_sizing.py`

- [x] 先编写失败测试，覆盖日期范围、行列范围、非法输入、单位换算及后写规则覆盖。
- [x] 运行 `python3 -m unittest tests.test_sizing -v`，确认因模块缺失而失败。
- [x] 实现范围解析和 Excel 行高/列宽换算。
- [x] 重跑测试并确认通过。

### Task 2: 任意日期工作簿转换

**Files:**
- Create: `src/monthly_workbook/workbook.py`
- Create: `tests/test_workbook.py`

- [x] 先编写失败测试，验证任意起止日期、标题、首张期初公式、后续累计公式、清空输入、行列尺寸和文件名。
- [x] 运行 `python3 -m unittest tests.test_workbook -v`，确认因接口缺失而失败。
- [x] 从已验证的 Skill 脚本提取转换逻辑，并按设计实现任意日期与尺寸规则。
- [x] 重跑测试并确认通过。

### Task 3: 桌面界面与帮助文档

**Files:**
- Create: `src/monthly_workbook/app.py`
- Create: `src/monthly_workbook/__main__.py`
- Create: `src/monthly_workbook/help.html`
- Create: `tests/test_app_helpers.py`

- [x] 先测试界面使用的日期解析、默认输出路径和规则数据转换。
- [x] 实现模板选择、标题、日期、输出位置、规则列表、生成及帮助入口。
- [x] 编写包含完整操作步骤和常见错误的离线 HTML 帮助。
- [x] 运行全部单元测试。

### Task 4: 双平台打包

**Files:**
- Create: `monthly-workbook.spec`
- Create: `scripts/build_macos.sh`
- Create: `scripts/build_windows.bat`
- Create: `.github/workflows/build.yml`
- Create: `README.md`

- [x] 配置 PyInstaller 收集 `help.html` 并分别生成 macOS `.app` 与 Windows `.exe`。
- [x] 添加本机构建脚本及 GitHub Actions 双平台构建任务。
- [x] 记录构建、运行和产物位置。
- [x] 在当前 macOS 环境安装构建依赖并实际构建 `.app`。

### Task 5: 端到端验证

**Files:**
- Create: `tests/inspect_output.py`

- [x] 用桌面上的真实模板生成跨月样例。
- [x] 检查页签、完整日期、标题、手工单元格、公式、行高和列宽。
- [x] 启动 macOS `.app` 并确认主窗口可打开。
- [x] 运行 `python3 -m unittest discover -s tests -v` 和构建检查，记录结果。
