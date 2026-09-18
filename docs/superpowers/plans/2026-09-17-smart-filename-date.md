# 文件名日期智能更新 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在用户修改起始或结束日期时，自动更新输出文件名的日期区间并保留自定义名称。

**Architecture:** 在 `app.py` 增加一个无 UI 依赖的文件名替换函数，由 Tkinter 日期变量的 trace 回调调用。纯函数单独测试，UI 只负责在日期有效时更新字段。

**Tech Stack:** Python 3.14, Tkinter, unittest, PyInstaller

---

### Task 1: 文件名日期替换

**Files:**
- Modify: `tests/test_app_helpers.py`
- Modify: `src/monthly_workbook/app.py`

- [x] **Step 1: Write the failing tests**

测试旧日期替换、自定义名称保留和无日期名称追加。

- [x] **Step 2: Run the focused tests and verify failure**

Run: `PYTHONPATH=src python3 -m unittest tests.test_app_helpers -v`
Expected: FAIL because `update_filename_dates` does not exist.

- [x] **Step 3: Implement the pure helper**

识别 `.xlsx` 前的 `YYYY-MM-DD-YYYY-MM-DD` 区间；有则替换，无则追加，并保留扩展名。

- [x] **Step 4: Connect date inputs**

为 `self.start` 和 `self.end` 注册 `trace_add("write", ...)`；回调只在模板已选择、两个日期都合法且结束日不早于起始日时更新文件名。

- [x] **Step 5: Run focused and full tests**

Run: `PYTHONPATH=src python3 -m unittest tests.test_app_helpers -v`
Run: `PYTHONPATH=src python3 -m unittest discover -s tests -v`
Expected: all tests pass.

### Task 2: 重建发布包

**Files:**
- Rebuild: `dist/月度生产表生成器.app`
- Replace: `release/月度生产表生成器-macOS-x86_64.zip`
- Replace: `release/月度生产表生成器-Windows构建包.zip`

- [x] **Step 1: Build macOS app**

Run: `sh scripts/build_macos.sh`
Expected: PyInstaller exits 0 and creates the `.app`.

- [x] **Step 2: Package both systems and copy to Desktop**

重建 macOS ZIP；将已更新源码、构建脚本和说明打入 Windows 构建包；复制两个 ZIP 到桌面。

- [x] **Step 3: Verify release artifacts**

Run ZIP integrity tests, `codesign --verify --deep --strict`, and confirm the Windows archive contains `update_filename_dates`.
