#!/usr/bin/env bash
# Renders the final decks to PDF + per-slide PNGs for visual review.
#
#   ./render.sh            # both languages
#   ./render.sh RU         # one language
#
# Requires LibreOffice (`soffice` on PATH) and poppler-utils (`pdftoppm`).
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(cd "$HERE/../.." && pwd)"
OUT="$ROOT/deck/render"

LANGS=("${@:-RU EN}")
read -r -a LANGS <<< "${LANGS[*]}"

SOFFICE="${SOFFICE:-soffice}"
command -v "$SOFFICE" >/dev/null || { echo "soffice not found; set SOFFICE=/path/to/soffice" >&2; exit 1; }

PROFILE="$(mktemp -d)"
mkdir -p "$OUT"

for L in "${LANGS[@]}"; do
  PPTX="$ROOT/deck/Efes_Nexus_$L.pptx"
  [ -f "$PPTX" ] || { echo "missing $PPTX (run: python3 $HERE/transform.py)" >&2; exit 1; }
  "$SOFFICE" --headless -env:UserInstallation="file://$PROFILE" \
             --convert-to pdf --outdir "$OUT" "$PPTX" >/dev/null
  pdftoppm -png -r 110 "$OUT/Efes_Nexus_$L.pdf" "$OUT/${L,,}_slide"
  echo "rendered $L -> $OUT"
done

rm -rf "$PROFILE"
ls -1 "$OUT"
