#!/usr/bin/env bash
# Font-independent copies of the lecture: a PDF (fonts embedded) and a picture-only PowerPoint.
# Needs LibreOffice, poppler (pdftoppm) and python-pptx; Arabic is rendered with Noto Sans Arabic.
set -euo pipefail
cd "$(dirname "$0")/.."
SRC="محاضرة_الكسوف_الكلي_2027.pptx"
TMP=$(mktemp -d)
cp "$SRC" "$TMP/deck.pptx"
soffice --headless --convert-to pdf --outdir "$TMP" "$TMP/deck.pptx" >/dev/null
cp "$TMP/deck.pdf" "محاضرة_الكسوف_الكلي_2027.pdf"
mkdir "$TMP/png" && pdftoppm -r 200 -png "$TMP/deck.pdf" "$TMP/png/s"
python3 -c "
import json, sys
from pptx import Presentation
p = Presentation(sys.argv[1])
json.dump([s.notes_slide.notes_text_frame.text if s.has_notes_slide else '' for s in p.slides],
          open(sys.argv[2], 'w', encoding='utf-8'), ensure_ascii=False)" "$SRC" "$TMP/notes.json"
node build/images.js "$TMP/png" "$TMP/notes.json" "محاضرة_الكسوف_الكلي_2027_مصورة.pptx"
rm -rf "$TMP"
