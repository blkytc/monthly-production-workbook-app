# Windows XLS Conversion And Theme Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add automatic Windows `.xls` conversion through installed Excel/WPS and apply a consistent dark application theme.

**Architecture:** Keep `legacy_reports.py` responsible for choosing a converter and add one packaged PowerShell script for Windows COM automation. Keep theme configuration in `app.py` as a small function invoked before the widgets are built.

**Tech Stack:** Python standard library, Tkinter/ttk, PowerShell COM automation, PyInstaller, unittest, GitHub Actions.

---

### Task 1: Windows XLS converter fallback

**Files:**
- Modify: `tests/test_legacy_reports.py`
- Modify: `src/monthly_workbook/legacy_reports.py`
- Create: `src/monthly_workbook/convert_xls_windows.ps1`

- [ ] Add tests that patch platform detection, converter discovery, and `subprocess.run`, then assert Windows invokes the packaged script and reports a useful error when no converter succeeds.
- [ ] Run `PYTHONPATH=src python3 -m unittest tests.test_legacy_reports.WindowsConversionTests -v` and confirm the new tests fail because the fallback does not exist.
- [ ] Add `_libreoffice_path`, `_windows_conversion_script`, and the minimal Windows PowerShell fallback in `convert_xls`.
- [ ] Implement the script to try `Excel.Application`, `ket.Application`, and `et.Application`, save with XLSX format 51, close the workbook, and quit the application in `finally` blocks.
- [ ] Run the focused tests and confirm they pass.

### Task 2: Portable dark ttk theme

**Files:**
- Modify: `tests/test_app_helpers.py`
- Modify: `src/monthly_workbook/app.py`

- [ ] Add a fake-style unit test that records `theme_use`, `configure`, and `map` calls and asserts portable `clam`, dark frames/entries/tree views, and selected states are configured.
- [ ] Run `PYTHONPATH=src python3 -m unittest tests.test_app_helpers.AppThemeTests -v` and confirm it fails because `apply_app_theme` does not exist.
- [ ] Implement `apply_app_theme(style, root)` with the approved palette and call it from `main` before constructing `WorkbookApp`.
- [ ] Run the focused theme tests and confirm they pass.

### Task 3: Package, document, and release

**Files:**
- Modify: `monthly-workbook.spec`
- Modify: `README.md`
- Modify: `src/monthly_workbook/help.html`

- [ ] Package `convert_xls_windows.ps1` in PyInstaller data and update user documentation to describe automatic installed-office conversion.
- [ ] Run `PYTHONPATH=src python3 -m unittest discover -s tests -v` and confirm all tests pass.
- [ ] Build the macOS app locally and launch it briefly to catch Tk startup or resource errors.
- [ ] Commit and push the implementation, tag `v1.0.3`, and wait for both GitHub Actions builds and the release job.
- [ ] Download both release ZIPs, verify archive integrity, confirm the Windows ZIP contains only `MonthlyWorkbookGenerator.exe`, and place the verified Windows ZIP on the desktop.
