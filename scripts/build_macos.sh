#!/bin/sh
set -eu
cd "$(dirname "$0")/.."
python3 -m PyInstaller --noconfirm --clean monthly-workbook.spec
echo "已生成 dist/月度生产表生成器.app"
