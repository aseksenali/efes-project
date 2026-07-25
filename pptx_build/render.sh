#!/usr/bin/env bash
set -e
cd /Users/asetseksenali/Downloads/efes
export PATH="/opt/homebrew/bin:$PATH"
SOFFICE=/Applications/LibreOffice.app/Contents/MacOS/soffice
rm -f Efes_Nexus.pdf
rm -f slide-*.jpg
"$SOFFICE" --headless --convert-to pdf --outdir . Efes_Nexus.pptx >/tmp/lo_conv.log 2>&1
echo "PDF ok"
pdftoppm -jpeg -r 130 Efes_Nexus.pdf slide
ls -1 "$PWD"/slide-*.jpg
