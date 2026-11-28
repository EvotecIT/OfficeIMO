"""Check that the format quick starts saved the files their READMEs describe."""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path

PACKAGES = ["Flow.vsdx", "Summary.odt", "Review.key", "Field-guide.epub"]
FILES = {
    "Delivery.xml": b"<Project",
    "People.csv": b"Name;Age;City;Bucket",
    "Report.rtf": b"{\\rtf1",
}


def verify(directory: Path) -> None:
    for name in PACKAGES:
        path = directory / name
        if not zipfile.is_zipfile(path):
            raise ValueError(f"{name} is not a ZIP package")
        with zipfile.ZipFile(path) as package:
            if package.testzip() is not None or not package.namelist():
                raise ValueError(f"{name} is empty or damaged")
    for name, marker in FILES.items():
        data = (directory / name).read_bytes()
        if marker not in data[:4096]:
            raise ValueError(f"{name} does not start like the expected format")
    for name in ["Contacts.accdb", "Planning.one"]:
        if (directory / name).stat().st_size == 0:
            raise ValueError(f"{name} is empty")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    verify(parser.parse_args().directory)
    print("Format quick starts saved their files.")
