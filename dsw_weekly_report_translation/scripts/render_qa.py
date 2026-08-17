#!/usr/bin/env python3
"""Render every slide of a .pptx to individual JPG images for visual QA.

Usage:
    python render_qa.py <input.pptx> [output_dir] [--dpi 150]

Requires LibreOffice (soffice) and pdftoppm. On macOS after
`brew install --cask libreoffice`, soffice is found at
/Applications/LibreOffice.app/Contents/MacOS/soffice.
Produces output_dir/slide-1.jpg ... slide-N.jpg.
"""
import shutil
import subprocess
import sys
from pathlib import Path


def find_soffice() -> str:
    for cand in (
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/local/bin/soffice",
        "/opt/homebrew/bin/soffice",
    ):
        if Path(cand).exists():
            return cand
    return shutil.which("soffice") or "soffice"


def render(pptx: str, outdir: str = ".", dpi: int = 150) -> list[str]:
    pptx_path = Path(pptx).resolve()
    out = Path(outdir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    soffice = find_soffice()
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out), str(pptx_path)],
        check=True,
        capture_output=True,
        text=True,
    )
    pdf = out / (pptx_path.stem + ".pdf")
    if not pdf.exists():
        raise FileNotFoundError(f"PDF conversion failed: {pdf}")

    subprocess.run(
        ["pdftoppm", "-jpeg", "-r", str(dpi), str(pdf), str(out / "slide")],
        check=True,
        capture_output=True,
        text=True,
    )
    return sorted(str(p) for p in out.glob("slide-*.jpg"))


if __name__ == "__main__":
    dpi = 150
    args = []
    for a in sys.argv[1:]:
        if a.startswith("--dpi"):
            dpi = int(a.split("=")[-1]) if "=" in a else 150
        else:
            args.append(a)
    if not args:
        print(__doc__)
        sys.exit(1)
    files = render(args[0], args[1] if len(args) > 1 else ".", dpi)
    print("\n".join(files))
