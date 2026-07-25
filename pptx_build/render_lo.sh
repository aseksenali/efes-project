#!/usr/bin/env bash
# Mount the downloaded LibreOffice dmg (no install / no brew) and render the pptx to per-slide JPGs.
set -e
cd /Users/asetseksenali/Downloads/efes
export PATH="/opt/homebrew/bin:$PATH"

DMG=/tmp/LibreOffice.dmg
MNT=/tmp/lo_mnt
SOFFICE=""

# detach any stale mount
hdiutil detach "$MNT" >/dev/null 2>&1 || true
rm -rf "$MNT"; mkdir -p "$MNT"
hdiutil attach "$DMG" -nobrowse -quiet -mountpoint "$MNT"
SOFFICE="$MNT/LibreOffice.app/Contents/MacOS/soffice"
[ -x "$SOFFICE" ] || SOFFICE="$(find "$MNT" -name soffice -path '*MacOS*' | head -1)"
echo "soffice: $SOFFICE"

PROFILE=/tmp/lo_profile
rm -rf "$PROFILE"; mkdir -p "$PROFILE"

rm -f Efes_Nexus.pdf
rm -f slide-*.jpg
"$SOFFICE" --headless -env:UserInstallation=file://$PROFILE \
  --convert-to pdf --outdir . Efes_Nexus.pptx
echo "PDF: $(ls -la Efes_Nexus.pdf 2>&1)"
pdftoppm -jpeg -r 130 Efes_Nexus.pdf slide
ls -1 "$PWD"/slide-*.jpg
hdiutil detach "$MNT" >/dev/null 2>&1 || true
