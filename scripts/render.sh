#!/usr/bin/env bash
# Render a deck to one JPG per slide for visual QA.
# usage: bash render.sh deck.pptx [dpi]
set -e
DECK="$1"; DPI="${2:-80}"
[ -f "$DECK" ] || { echo "usage: render.sh deck.pptx [dpi]"; exit 2; }
command -v soffice >/dev/null 2>&1 || command -v libreoffice >/dev/null 2>&1 || { echo "LibreOffice (soffice) not found - install it to render previews"; exit 3; }
command -v pdftoppm >/dev/null 2>&1 || { echo "pdftoppm not found - install poppler / poppler-utils"; exit 3; }
SOFFICE=$(command -v soffice || command -v libreoffice)
DIR=$(cd "$(dirname "$DECK")" && pwd); BASE=$(basename "${DECK%.*}")
"$SOFFICE" --headless --convert-to pdf --outdir "$DIR" "$DECK" >/dev/null 2>&1
rm -f "$DIR/$BASE"-slide-*.jpg
pdftoppm -jpeg -r "$DPI" "$DIR/$BASE.pdf" "$DIR/$BASE-slide"
ls "$DIR/$BASE"-slide-*.jpg
