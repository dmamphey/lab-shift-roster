#!/usr/bin/env python3
"""Render the user guide to PDF.

    python tools/build_guide_pdf.py

The PDF is produced by printing ``user-guide.html`` from a headless browser, so
there is one source for the guide and the two versions cannot drift apart. The
print stylesheet in the page keeps the brand colours and the contents list, and
because a real browser engine does the rendering, every link in the page survives
as a clickable annotation in the PDF.

Chrome or Edge must be installed; nothing is downloaded and nothing is uploaded.

Relative links (``index.html``) are pointed at the published site in a temporary
copy of the page: a relative href means nothing once the file is a PDF sitting in
someone's downloads folder, but rewriting the page itself would break the guide
for anyone running the app locally.
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "user-guide.html"
DEFAULT_OUTPUT = ROOT / "Lab-Shift-Roster-1.1-User-Guide.pdf"

SITE = "https://dmamphey.github.io/lab-shift-roster/"

#: Where Chrome and Edge install themselves on Windows, then the usual names on
#: PATH for macOS and Linux.  The first one that exists is used.
CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "google-chrome",
    "chromium",
    "microsoft-edge",
]


def find_browser() -> str | None:
    for candidate in CANDIDATES:
        if Path(candidate).is_file():
            return candidate
        found = shutil.which(candidate)
        if found:
            return found
    return None


def prepare_page(html: str) -> str:
    """Point relative links at the published site."""
    return re.sub(r'href="(?!#|https?:|data:|mailto:)([^"]+)"',
                  lambda match: f'href="{SITE}{match.group(1)}"', html)


def build(browser: str, output: Path) -> int:
    # A working directory without spaces keeps the file:// URL and the command
    # line straightforward on Windows.
    with tempfile.TemporaryDirectory(prefix="labguide") as tmp:
        work = Path(tmp)
        page = work / "guide.html"
        page.write_text(prepare_page(SOURCE.read_text(encoding="utf-8")),
                        encoding="utf-8")
        pdf = work / "guide.pdf"

        command = [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--no-first-run",
            "--no-default-browser-check",
            "--force-color-profile=srgb",
            # The page is static, but give web fonts and layout a moment to
            # settle before the snapshot is taken.
            "--virtual-time-budget=4000",
            "--no-pdf-header-footer",
            f"--print-to-pdf={pdf}",
            page.as_uri(),
        ]
        result = subprocess.run(command, capture_output=True, text=True,
                                timeout=180)
        if not pdf.exists():
            print(f"{Path(browser).name} produced no PDF "
                  f"(exit {result.returncode})")
            print(result.stderr.strip()[:2000])
            return 1

        output.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(pdf, output)

    size = output.stat().st_size
    print(f"Rendered with {Path(browser).name}")
    print(f"Wrote {output.relative_to(ROOT) if output.is_relative_to(ROOT) else output}"
          f" ({size / 1024:.0f} KB)")
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Render the Lab Shift Roster user guide to PDF.")
    parser.add_argument("--out", default=str(DEFAULT_OUTPUT),
                        help="output path (default: %(default)s)")
    parser.add_argument("--browser", default=None,
                        help="path to a Chrome or Edge executable")
    args = parser.parse_args(argv)

    if not SOURCE.exists():
        print(f"Cannot find {SOURCE}")
        return 1

    browser = args.browser or find_browser()
    if not browser:
        print("No Chrome or Edge installation found. Install either one, or "
              "pass --browser with the path to its executable.")
        return 1

    return build(browser, Path(args.out).resolve())


if __name__ == "__main__":
    sys.exit(main())
