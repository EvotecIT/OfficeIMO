"""Check the saved Markdown and HTML quick starts."""

from __future__ import annotations

import argparse
from pathlib import Path


def verify(directory: Path) -> None:
    markdown = (directory / "Release-notes.md").read_text(encoding="utf-8")
    if not markdown.startswith("---") or "# Release notes" not in markdown or "- Parse Markdown into typed blocks and inlines" not in markdown:
        raise ValueError("The Markdown quick start is missing its front matter, heading or list")
    html = (directory / "Release-notes.html").read_text(encoding="utf-8")
    if "<h1" not in html or "Release notes" not in html or "<li>" not in html:
        raise ValueError("The Markdown quick start did not render its HTML document")
    report = (directory / "Service-review.html").read_text(encoding="utf-8")
    if "Service review" not in report or "<td>Certificates</td><td>Renewed</td>" not in report:
        raise ValueError("The HTML quick start is missing its edited table row")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    verify(parser.parse_args().directory)
    print("Current-source Markdown and HTML quick starts passed.")
