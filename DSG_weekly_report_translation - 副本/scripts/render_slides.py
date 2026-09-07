#!/usr/bin/env python3
"""Render every slide of a .pptx to individual JPG images for visual QA.

Usage:
    python render_slides.py <input.pptx> [output_dir] [--dpi 150]

Produces output_dir/slide-1.jpg, slide-2.jpg, ... Requires LibreOffice
(soffice) and pdftoppm, both available in the Claude container. Use this
after every editing pass to visually compare against the source deck and
the template (references/template.pptx).
"""
import subprocess
import sys
from pathlib import Path


def render(pptx: str, outdir: str = ".", dpi: int = 150) -> list[str]:
    pptx_path = Path(pptx).resolve()
    out = Path(outdir).resolve()
    out.mkdir(parents=True, exist_ok=True)

    # pptx -> pdf (soffice.py wrapper handles profile isolation if present)
    soffice_wrapper = Path("/mnt/skills/public/pptx/scripts/office/soffice.py")
    if soffice_wrapper.exists():
        cmd = [sys.executable, str(soffice_wrapper), "--headless",
               "--convert-to", "pdf", "--outdir", str(out), str(pptx_path)]
    else:
        cmd = ["soffice", "--headless", "--convert-to", "pdf",
               "--outdir", str(out), str(pptx_path)]
    subprocess.run(cmd, check=True, capture_output=True, text=True)

    pdf = out / (pptx_path.stem + ".pdf")
    if not pdf.exists():
        raise FileNotFoundError(f"PDF conversion failed: {pdf}")

    # pdf -> per-page jpgs
    subprocess.run(["pdftoppm", "-jpeg", "-r", str(dpi), str(pdf),
                    str(out / "slide")], check=True)
    return sorted(str(p) for p in out.glob("slide-*.jpg"))


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dpi = 150
    for a in sys.argv[1:]:
        if a.startswith("--dpi"):
            dpi = int(a.split("=")[-1]) if "=" in a else 150
    if not args:
        print(__doc__)
        sys.exit(1)
    files = render(args[0], args[1] if len(args) > 1 else ".", dpi)
    print("\n".join(files))
