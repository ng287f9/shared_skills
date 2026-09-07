#!/usr/bin/env python
"""ASCII downsampled view of a screenshot to sanity-check table structure.
Usage: python3 ascii_view.py <image.png> [cols=110] [rows=40]
"""
import sys
from PIL import Image

def main(img, cols=110, rows=40):
    im = Image.open(img).convert('L').resize((cols, rows))
    px = im.load()
    for y in range(rows):
        print(''.join('#' if px[x, y] < 128 else ('+' if px[x, y] < 200 else '.') for x in range(cols)))

if __name__ == '__main__':
    main(sys.argv[1], int(sys.argv[2]) if len(sys.argv) > 2 else 110, int(sys.argv[3]) if len(sys.argv) > 3 else 40)
