#!/usr/bin/env python3
"""Download the Pyodide runtime and Python wheels for self-hosted deployment.

Run this once before deploying to a site that must serve everything from a single
trusted domain, which is the usual requirement in an NHS-managed environment:

    python tools/fetch_runtime.py

It writes about 14 MB into ``vendor/``, which is deliberately excluded from git.
The browser prefers ``vendor/`` when it is present and falls back to the public CDN
when it is not, so the application works either way and development needs no
download at all.

Licences of what is fetched — all permit redistribution, and the notices are
written into vendor/LICENCES.txt alongside the files:

    Pyodide      Mozilla Public License 2.0
    CPython      Python Software Foundation License 2.0
    micropip     Mozilla Public License 2.0
    packaging    Apache 2.0 / BSD dual
    openpyxl     MIT
    et_xmlfile   MIT
"""

from __future__ import annotations

import argparse
import sys
import urllib.error
import urllib.request
from pathlib import Path

PYODIDE_VERSION = "0.26.4"
CDN = f"https://cdn.jsdelivr.net/pyodide/v{PYODIDE_VERSION}/full"

#: The runtime files the browser actually loads.  Deliberately not the whole
#: Pyodide distribution, which is several hundred megabytes of packages LabRoster
#: never touches.
RUNTIME_FILES = [
    "pyodide.js",
    "pyodide.mjs",
    "pyodide.asm.js",
    "pyodide.asm.wasm",
    "python_stdlib.zip",
    "pyodide-lock.json",
]

#: Wheels Pyodide resolves through its own lock file.  These must sit beside the
#: runtime, because loadPackage() looks them up relative to indexURL.
RUNTIME_WHEELS = [
    "micropip-0.6.0-py3-none-any.whl",
    "packaging-23.2-py3-none-any.whl",
]

#: Wheels LabRoster installs by URL, which micropip would otherwise pull from
#: PyPI at run time.
WHEELS = [
    "https://files.pythonhosted.org/packages/py2.py3/o/openpyxl/"
    "openpyxl-3.1.5-py2.py3-none-any.whl",
    "https://files.pythonhosted.org/packages/py3/e/et_xmlfile/"
    "et_xmlfile-2.0.0-py3-none-any.whl",
]

LICENCES = """LabRoster self-hosted runtime — third party notices

Pyodide (runtime, micropip)      Mozilla Public License 2.0
    https://github.com/pyodide/pyodide/blob/main/LICENSE
CPython standard library         Python Software Foundation License 2.0
    https://docs.python.org/3/license.html
packaging                        Apache 2.0 / BSD (dual)
    https://github.com/pypa/packaging/blob/main/LICENSE
openpyxl                         MIT
    https://foss.heptapod.net/openpyxl/openpyxl
et_xmlfile                       MIT
    https://foss.heptapod.net/openpyxl/et_xmlfile

All of the above permit redistribution.  Keep this file alongside the runtime so
the notices travel with the files they describe.
"""


def download(url: str, target: Path, force: bool) -> int:
    if target.exists() and not force:
        print(f"  have  {target.name} ({target.stat().st_size:,} bytes)")
        return target.stat().st_size
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        with urllib.request.urlopen(url, timeout=120) as response:
            data = response.read()
    except (urllib.error.URLError, TimeoutError) as error:
        print(f"  FAILED {target.name}: {error}")
        return 0
    target.write_bytes(data)
    print(f"  got   {target.name} ({len(data):,} bytes)")
    return len(data)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Fetch the Pyodide runtime and wheels for self-hosting.")
    parser.add_argument("--into", default="vendor",
                        help="directory to write into (default: vendor)")
    parser.add_argument("--force", action="store_true",
                        help="re-download files that are already present")
    args = parser.parse_args(argv)

    root = Path(args.into)
    total = 0

    print(f"Pyodide {PYODIDE_VERSION} runtime -> {root / 'pyodide'}")
    for name in RUNTIME_FILES + RUNTIME_WHEELS:
        total += download(f"{CDN}/{name}", root / "pyodide" / name, args.force)

    print(f"\nPython wheels -> {root / 'wheels'}")
    for url in WHEELS:
        total += download(url, root / "wheels" / url.rsplit("/", 1)[-1],
                          args.force)

    (root / "LICENCES.txt").write_text(LICENCES, encoding="utf-8")
    print(f"\nWrote third party notices to {root / 'LICENCES.txt'}")
    print(f"Total: {total / 1048576:.2f} MB")

    if total == 0:
        print("\nNothing was downloaded. Check network access to "
              "cdn.jsdelivr.net and files.pythonhosted.org.")
        return 1

    print("\nDeploy the whole folder, including vendor/. The page will detect "
          "vendor/pyodide and stop using the public CDN.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
