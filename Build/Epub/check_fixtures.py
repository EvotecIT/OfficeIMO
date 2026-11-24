#!/usr/bin/env python3
"""Generate OfficeIMO publications and check them with the pinned EPUBCheck distribution."""
import argparse
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from urllib.request import urlopen
from zipfile import ZipFile

from validate_epub import digest


VERSION = "5.4.0"
SHA256 = "33350c61038e71dfb3d45a76aed04bf5481e6d5500cb780f6e98db8bbd15a28c"
DISTRIBUTION = f"epubcheck-{VERSION}"
URL = f"https://github.com/w3c/epubcheck/releases/download/v{VERSION}/{DISTRIBUTION}.zip"
# These fixtures deliberately preserve caller-supplied CSS rejected by EPUBCheck.
PRESERVATION_CASES = {"css-preservation.epub", "merge-nested-selectors.epub"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path, help="New task-owned evidence directory")
    parser.add_argument("--java", default="java")
    parser.add_argument("--include-preservation-cases", action="store_true",
                        help="Also validate the documented CSS preservation cases, which fail EPUBCheck")
    parser.add_argument("--ace", type=Path, help="Optional installed Ace CLI executable")
    parser.add_argument("epubs", nargs="*", type=Path, help="Additional publications to validate")
    args = parser.parse_args()

    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    scripts = Path(__file__).resolve().parent
    fixtures = output / "fixtures"
    with tempfile.TemporaryDirectory(prefix="officeimo-epubcheck-") as temp:
        temp = Path(temp)
        archive = temp / f"{DISTRIBUTION}.zip"
        with urlopen(URL, timeout=60) as source, archive.open("xb") as target:
            shutil.copyfileobj(source, target)
        if digest(archive) != SHA256:
            raise ValueError(f"SHA-256 mismatch for {DISTRIBUTION}.zip")
        (output / "epubcheck-distribution.txt").write_text(f"{SHA256}  {DISTRIBUTION}.zip\n{URL}\n", encoding="utf-8")
        with ZipFile(archive) as package:
            package.extractall(temp)
        subprocess.run(["dotnet", "run", "--configuration", "Release", "--project",
                        str(scripts / "Fixtures"), "--", str(fixtures)],
                       cwd=scripts.parents[1], check=True)

        generated = sorted(fixtures.glob("*.epub"))
        if not args.include_preservation_cases:
            excluded = sorted(path.name for path in generated if path.name in PRESERVATION_CASES)
            generated = [path for path in generated if path.name not in PRESERVATION_CASES]
            if excluded:
                print("CSS preservation cases excluded: " + ", ".join(excluded), flush=True)
        if not generated:
            raise ValueError("The fixture generator did not produce any publications")
        command = [sys.executable, str(scripts / "validate_epub.py"),
                   "--epubcheck-jar", str(temp / DISTRIBUTION / "epubcheck.jar"),
                   "--java", args.java, "--output", str(output / "reports")]
        if args.ace:
            command.extend(("--ace", str(args.ace)))
        command.extend(("--", *(str(path) for path in [*generated, *args.epubs])))
        # EPUBCheck inherits message-severity overrides from this variable.
        environment = os.environ.copy()
        environment.pop("ePubCheckCustomMessageFile", None)
        return subprocess.run(command, env=environment, check=False).returncode


if __name__ == "__main__":
    sys.exit(main())
