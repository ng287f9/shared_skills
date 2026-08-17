#!/usr/bin/env python3
"""
Convenience runner: execute all three phases of FT_Plan_Update in sequence.

Usage:
    python scripts/run_all.py

Or with custom paths:
    python scripts/run_all.py --mpp FlightTestSchedule_20260626_PlanB.mpp --cn 术语对照_中文.xlsx --en 术语对照_英文.xlsx

All scripts are looked up relative to this script's directory.
"""

import argparse
import os
import shutil
import subprocess
import sys
from datetime import datetime


def copy_unlocked(src, dst):
    """Copy a file, working around OneDrive sync locks."""
    for attempt in range(3):
        try:
            shutil.copy2(src, dst)
            return True
        except PermissionError:
            if attempt < 2:
                import time
                time.sleep(2)
    return False


def get_script_dir():
    return os.path.dirname(os.path.abspath(__file__))


def find_mpp_file(cwd):
    """Find the first FlightTestSchedule_*.mpp file in cwd."""
    for f in sorted(os.listdir(cwd)):
        if f.startswith("FlightTestSchedule_") and f.endswith(".mpp"):
            return f
    return None


def main():
    script_dir = get_script_dir()
    cwd = os.getcwd()

    parser = argparse.ArgumentParser(description="Run all FT_Plan_Update phases")
    parser.add_argument("--mpp", default=None, help="Path to .mpp file (default: auto-detect)")
    parser.add_argument("--cn", default="术语对照_中文.xlsx", help="Chinese terminology Excel")
    parser.add_argument("--en", default="术语对照_英文.xlsx", help="English terminology Excel")
    args = parser.parse_args()

    # Resolve paths
    mpp_path = args.mpp or find_mpp_file(cwd)
    if not mpp_path:
        print("Error: No .mpp file found. Specify with --mpp", file=sys.stderr)
        sys.exit(1)

    cn_path = os.path.join(cwd, args.cn) if not os.path.isabs(args.cn) else args.cn
    en_path = os.path.join(cwd, args.en) if not os.path.isabs(args.en) else args.en
    mpp_full = os.path.join(cwd, mpp_path) if not os.path.isabs(mpp_path) else mpp_path

    # Check input files
    for label, path in [("CN", cn_path), ("EN", en_path), ("MPP", mpp_full)]:
        if not os.path.exists(path):
            print(f"Error: {label} file not found: {path}", file=sys.stderr)
            sys.exit(1)

    # Output paths
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # Include seconds to avoid conflicts
    mpp_stem = os.path.splitext(os.path.basename(mpp_path))[0]
    glossary_path = os.path.join(cwd, "references", "术语对照.md")
    excel_output = os.path.join(cwd, f"FT_Plan_Update_{mpp_stem}_{timestamp}.xlsx")

    os.makedirs(os.path.dirname(glossary_path), exist_ok=True)

    # Phase 1: Build glossary (work around OneDrive locks)
    print("=" * 60)
    print("PHASE 1: Building terminology glossary")
    print("=" * 60)
    phase1 = os.path.join(script_dir, "build_glossary.py")

    # Copy locked files to temp
    cn_tmp = cn_path + ".tmp_for_read.xlsx"
    en_tmp = en_path + ".tmp_for_read.xlsx"
    cn_ok = copy_unlocked(cn_path, cn_tmp)
    en_ok = copy_unlocked(en_path, en_tmp)

    result = subprocess.run(
        [sys.executable, phase1,
         "--cn", cn_tmp if cn_ok else cn_path,
         "--en", en_tmp if en_ok else en_path,
         "--output", glossary_path],
        cwd=cwd,
    )
    # Cleanup temps
    if cn_ok and os.path.exists(cn_tmp): os.remove(cn_tmp)
    if en_ok and os.path.exists(en_tmp): os.remove(en_tmp)

    if result.returncode != 0:
        print("Phase 1 failed!", file=sys.stderr)
        sys.exit(1)

    # Phase 2: MPP → Excel (build from MPP with MPP formatting)
    print()
    print("=" * 60)
    print("PHASE 2: MPP → Excel (build from MPP data, match MPP formatting)")
    print("=" * 60)
    phase2 = os.path.join(script_dir, "read_mpp_to_excel.py")

    # Copy template to avoid OneDrive lock
    tmpl_tmp = cn_path + ".tmp_template.xlsx"
    tmpl_ok = copy_unlocked(cn_path, tmpl_tmp)

    result = subprocess.run(
        [sys.executable, phase2,
         "--mpp", mpp_full,
         "--template", tmpl_tmp if tmpl_ok else cn_path,
         "--output", excel_output],
        cwd=cwd,
    )
    if tmpl_ok and os.path.exists(tmpl_tmp): os.remove(tmpl_tmp)

    if result.returncode != 0:
        print("Phase 2 failed!", file=sys.stderr)
        sys.exit(1)

    # Phase 3: Translate (English → Chinese using glossary)
    print()
    print("=" * 60)
    print("PHASE 3: Translating English task names to Chinese")
    print("=" * 60)
    phase3 = os.path.join(script_dir, "translate_excel.py")
    result = subprocess.run(
        [sys.executable, phase3, "--input", excel_output, "--glossary", glossary_path, "--output", excel_output],
        cwd=cwd,
    )
    if result.returncode != 0:
        print("Phase 3 failed!", file=sys.stderr)
        sys.exit(1)

    # Phase 4 (Optional): Compare with previous version and highlight changes
    print()
    print("=" * 60)
    print("PHASE 4: Comparing with previous version and highlighting changes")
    print("=" * 60)

    prev_file = os.path.join(cwd, "previous.xlsx")
    if os.path.exists(prev_file):
        phase4 = os.path.join(script_dir, "compare_and_highlight.py")
        result = subprocess.run(
            [sys.executable, phase4,
             "--new", excel_output,
             "--previous", prev_file,
             "--output", excel_output],
            cwd=cwd,
        )
        if result.returncode != 0:
            print("Phase 4 (comparison) completed with warnings", file=sys.stderr)
    else:
        print("No previous.xlsx found. Skipping comparison phase.")

    print()
    print("=" * 60)
    print(f"ALL PHASES COMPLETE")
    print(f"Output: {excel_output}")
    print(f"Glossary: {glossary_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
