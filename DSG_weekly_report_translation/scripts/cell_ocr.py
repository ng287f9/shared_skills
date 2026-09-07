#!/usr/bin/env python
"""Per-cell OCR of a table screenshot. Crops each cell from h/v line midpoints,
upscales + binarizes, runs tesseract. NOTE: write temp crops to a local relative
dir (ocr_tmp/) — macOS tesseract fails on absolute /tmp paths.

Usage: python3 cell_ocr.py <image.png> "<rows list>" "<cols list>"
Example: python3 cell_ocr.py ppt/media/image7.png "[13,78,163,229,294,360]" "[6,254,465,677,904]"
"""
import os, subprocess, sys
from PIL import Image, ImageOps

def main(img, rows, cols):
    rows = eval(rows); cols = eval(cols)
    os.makedirs('ocr_tmp', exist_ok=True)
    im = Image.open(img).convert('L')
    for r in range(len(rows) - 1):
        line = []
        for c in range(len(cols) - 1):
            box = (cols[c] + 2, rows[r] + 2, cols[c + 1] - 2, rows[r + 1] - 2)
            crop = im.crop(box)
            crop = ImageOps.autocontrast(crop.resize((crop.width * 4, crop.height * 4), Image.LANCZOS))
            crop = crop.point(lambda p: 255 if p > 140 else 0)
            tmp = os.path.join('ocr_tmp', 'cell.png')
            crop.save(tmp)
            out = subprocess.run(['tesseract', tmp, 'ocr_tmp/cell', '--psm', '6'],
                                 capture_output=True, text=True)
            txt = ''
            if os.path.exists('ocr_tmp/cell.txt'):
                txt = open('ocr_tmp/cell.txt').read().strip().replace('\n', ' ')
            line.append(txt)
        print(f"row{r}:", " | ".join(line))

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3])
