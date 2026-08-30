#!/usr/bin/env python3
"""Classify MISC folder files by filename prefix + title heuristics.

Rules:
  AN<digit>*        -> Standards/AN
  MS<digit>* / MSM* -> Standards/MS
  AND<digit>*       -> Standards/AND
  ESA_PSS*          -> Standards/ESA
  ecss*             -> Standards/ECSS
  known agency code -> Reports/<CODE>
  otherwise         -> Books
  fake/broken files -> _Invalid

The matching .md (same stem) is moved alongside its PDF.
"""

import os
import re
import shutil
import sys

ROOT = r"F:\abbotspace\downloads\MISC"

AGENCIES = {
    "AEDC", "AFAMRL", "AFAPL", "AFATL", "AFCRL", "AFFDL", "AFFTC", "AFGM",
    "AFHRL", "AFIT", "AFML", "AFOSR", "AFRL", "AFRPL", "AFWAL", "AFWL",
    "AGATE", "ARC", "ARBRL", "ARL", "ARMY", "ASD", "ASTIA", "ATIC", "ATSB",
    "AVRADCOM", "AVSCOM", "BASI", "BIOS", "BTS", "BUAER", "CAA", "CERL",
    "CINDAS", "CONVAIR", "DAAC", "DARCOM", "DGCA", "DMIC", "DNA", "DOE",
    "DON", "DOT", "DREO", "DSTO", "DTIC", "EASA", "ESA", "ERDC", "FPL",
    "FTD", "GAO", "ICAS", "IDA", "KAPL", "MASAAG", "MCIC", "MDAC", "MDC",
    "MGU", "MTL", "NAA", "NADC", "NAEC", "NAIC", "NARCAP", "NATICK",
    "NAVAER", "NAVAIR", "NAVEDTRA", "NAVORD", "NAVROD", "NAVSEA", "NAVWEPS",
    "NAWCADPAX", "NBSHA", "NISTIR", "NRCC", "NRL", "NSMRL", "NSRDC", "NSWC",
    "NSWCDD", "NTIS", "NTSB", "NWC", "NZCAA", "ONR", "OTS", "RAE", "RAES",
    "RISO", "RTO", "SNL", "TPRC", "TREC", "TSARCOM", "USAAMRDL", "USAARL",
    "USAAVRADCOM", "USAAVSCOM", "USAF", "USAFA", "USAMC", "USARTL",
    "USASVSCOM", "USDA", "USDC", "WADC", "WADD", "WL", "WRDC", "AAMRL",
    "AMCP", "AMMME", "AMMRC", "AMRL", "AAIB", "AATB", "ACAAR", "ASME",
    "AWS", "CSDL", "ESL", "EHEST", "HFL", "ICAT", "ADPO10769",
}


def norm_prefix(stem: str) -> str:
    return stem.split("-")[0].strip()


def classify(stem: str) -> str:
    prefix = norm_prefix(stem)
    p = prefix.upper().replace(" ", "_")
    if re.match(r"^AND\d", p):
        return r"Standards\AND"
    if re.match(r"^AN_?\d", p):
        return r"Standards\AN"
    if re.match(r"^MSM?\d", p):
        return r"Standards\MS"
    if p.startswith("ESA_PSS"):
        return r"Standards\ESA"
    if prefix.lower() == "ecss" or p.startswith("ECSS"):
        return r"Standards\ECSS"
    for code in AGENCIES:
        if p == code or p.startswith(code + "_") or p.startswith(code + " "):
            return f"Reports\\{code}"
    if p.endswith(".EXE"):
        return "_Invalid"
    return "Books"


def main() -> int:
    files = [f for f in os.listdir(ROOT) if os.path.isfile(os.path.join(ROOT, f))]
    stems = {}
    plan = []
    conflicts = []
    counts: dict[str, int] = {}

    for f in sorted(files):
        stem, ext = os.path.splitext(f)
        ext = ext.lower()
        if ext == ".pdf":
            stems.setdefault(stem, {})["pdf"] = f
        elif ext == ".md":
            stems.setdefault(stem, {})["md"] = f

    for stem, group in stems.items():
        # classify by the pdf if present, else by the md (same stem)
        name = group.get("pdf") or group.get("md")
        dest = classify(os.path.splitext(name)[0])
        counts[dest] = counts.get(dest, 0) + 1
        for f in group.values():
            plan.append((f, dest))

    print("=== 归类计划 ===")
    for k in sorted(counts):
        print(f"{k}: {counts[k]} 组")

    moved = removed_dup = 0
    for f, dest in plan:
        src = os.path.join(ROOT, f)
        dst_dir = os.path.join(ROOT, dest)
        os.makedirs(dst_dir, exist_ok=True)
        dst = os.path.join(dst_dir, f)
        if os.path.exists(dst):
            # leftover from an interrupted run: copy exists, maybe src too
            if os.path.exists(src) and \
                    os.path.getsize(src) == os.path.getsize(dst):
                os.remove(src)
                removed_dup += 1
                continue
            conflicts.append(dst)
            continue
        if not os.path.exists(src):
            continue
        shutil.move(src, dst)
        moved += 1

    print(f"=== 已移动 {moved} 个, 清理中断副本 {removed_dup} 个 ===")
    if conflicts:
        print(f"冲突跳过 {len(conflicts)} 个:")
        for c in conflicts[:20]:
            print("  ", os.path.basename(c))
    return 0


if __name__ == "__main__":
    sys.exit(main())
