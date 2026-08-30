#!/usr/bin/env python3
"""Re-download broken PDFs from abbottaerospace.com technical library.

For each target: search site -> open /downloads/<slug>/ page -> extract
wpdm-filelist rows -> pick best-matching file -> download via proxy.
"""

import os
import re
import sys
import time
import urllib.parse
import urllib.request

PROXY = {"http": "http://127.0.0.1:10809", "https": "http://127.0.0.1:10809"}
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler(PROXY))
OPENER.addheaders = [("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")]

DEST = r"F:\abbotspace\downloads\MISC\_Invalid\redownloaded"

TARGETS = [
    ("AD-A951936 Thermophysical Properties of Matter - Vol. 2 - Thermal Conductivity.pdf",
     ["thermophysical properties of matter thermal conductivity"]),
    ("AD-A951939 Thermophysical Properties of Matter - Vol. 5 - Specific Heat.pdf",
     ["thermophysical properties of matter specific heat"]),
    ("AD-A951941 Thermophysical Properties of Matter - Vol. 7 - Thermal Radiative Properties.pdf",
     ["thermophysical properties of matter thermal radiative"]),
    ("ADS-71-SP Environmental Airworthiness and Qualification Requirements.pdf",
     ["ADS-71-SP environmental airworthiness"]),
    ("AEDC-TR-75-125 Static Stability and Fin Loads Data.pdf",
     ["AEDC-TR-75-125"]),
    ("AFWAL-TR-80-4110 Metrication of Mil-Hdbk-5C.pdf",
     ["AFWAL-TR-80-4110"]),
    ("NRL-RP-4276 U.A.R. Report No. XXI Upper Atmosphere Research Firings.pdf",
     ["NRL-RP-4276"]),
    ("RTO-SCI-040 Flight Test Measurement Techniques for Laminar Flow.pdf",
     ["RTO-SCI-040"]),
    ("XL Viking Display Math.xlam",
     ["XL Viking"]),
]


def fetch(url: str, timeout: int = 60) -> bytes:
    with OPENER.open(url, timeout=timeout) as resp:
        return resp.read()


def search_slugs(keywords: str) -> list[str]:
    q = urllib.parse.quote(keywords)
    html = fetch(f"https://www.abbottaerospace.com/?s={q}").decode("utf-8", "replace")
    return list(dict.fromkeys(re.findall(
        r'href="(https://www\.abbottaerospace\.com/downloads/[^"?#]+)"', html)))


def filelist(page_html: str) -> list[tuple[str, str]]:
    """Return [(filename, download_url)] from a download page."""
    out = []
    for m in re.finditer(
            r"<td>([^<]+\.pdf)</td><td[^>]*>\s*<a[^>]+href='([^']*wpdmdl=\d+[^']*)'",
            page_html, re.I):
        out.append((m.group(1).strip(), m.group(2)))
    return out


def score(fname: str, keywords: list[str]) -> int:
    f = fname.lower()
    return sum(1 for k in keywords.split() if k in f)


def main() -> int:
    os.makedirs(DEST, exist_ok=True)
    ok = fail = 0
    for final_name, queries in TARGETS:
        print(f"=== {final_name}", flush=True)
        got = False
        for q in queries:
            try:
                slugs = search_slugs(q)
            except Exception as exc:
                print(f"  search error: {exc}", flush=True)
                continue
            print(f"  [{q}] {len(slugs)} slugs", flush=True)
            for slug in slugs[:5]:
                try:
                    page = fetch(slug).decode("utf-8", "replace")
                except Exception as exc:
                    print(f"  page error {slug}: {exc}", flush=True)
                    continue
                files = filelist(page)
                if not files:
                    continue
                best = max(files, key=lambda x: score(x[0].lower(), q))
                fname, url = best
                print(f"  -> {slug.split('/')[-2]}: {fname}", flush=True)
                try:
                    data = fetch(url, timeout=300)
                    if data[:5] != b"%PDF-" and not final_name.endswith(".xlam"):
                        print(f"     not a PDF ({len(data)} bytes, "
                              f"head={data[:20]!r})", flush=True)
                        continue
                    path = os.path.join(DEST, fname)
                    with open(path, "wb") as fh:
                        fh.write(data)
                    print(f"     saved {len(data)//1024} KB -> {fname}", flush=True)
                    got = True
                    break
                except Exception as exc:
                    print(f"     download error: {exc}", flush=True)
            if got:
                break
        if got:
            ok += 1
        else:
            fail += 1
            print("  FAILED", flush=True)
        time.sleep(1)
    print(f"DONE ok={ok} fail={fail}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
