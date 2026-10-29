#!/usr/bin/env bash
# Збирає PDF + Excel, перераховує формули в Excel (LibreOffice) і перевіряє суми.
# Запускати з робочої папки, де лежить grili_data.py:  bash <шлях_до_скілла>/scripts/build.sh [вихідна_папка]
set -euo pipefail
S="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT="${1:-out}"
mkdir -p "$OUT/_rc"
export GRILI_OUT="$OUT"
python3 "$S/grili_make_pdf.py"
python3 "$S/grili_make_xlsx.py"
for f in "$OUT"/*.xlsx; do
  soffice --headless --convert-to xlsx --outdir "$OUT/_rc" "$f" >/dev/null 2>&1
  python3 "$S/verify_xlsx.py" "$OUT/_rc/$(basename "$f")"
  cp "$OUT/_rc/$(basename "$f")" "$f"   # версія з порахованими значеннями
done
rm -rf "$OUT/_rc" logo_xl.png
echo "Готово:"; ls -1 "$OUT"
