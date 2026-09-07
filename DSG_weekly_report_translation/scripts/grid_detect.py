#!/usr/bin/env python
"""Detect table grid lines in a screenshot; prints row/col line midpoints as lists.
Usage: python3 grid_detect.py <image.png>
"""
import sys
import numpy as np
from PIL import Image

def main(img):
    im = np.array(Image.open(img).convert('L'))
    h, w = im.shape
    dark = im < 128
    row_frac = dark.sum(axis=1) / w
    col_frac = dark.sum(axis=0) / h
    def midpoints(frac, thresh=0.5):
        idx = np.where(frac > thresh)[0]
        groups, cur = [], [idx[0]] if len(idx) else []
        for a, b in zip(idx, idx[1:]):
            if b - a <= 2: cur.append(b)
            else: groups.append(cur); cur = [b]
        if cur: groups.append(cur)
        return [int(np.mean(g)) for g in groups]
    print("rows:", midpoints(row_frac))
    print("cols:", midpoints(col_frac))

if __name__ == '__main__':
    main(sys.argv[1])
