# EPUB validation evidence

`validate_epub.py` runs an explicitly supplied EPUBCheck distribution over captured
publication bytes. It is an opt-in validation tool, outside normal builds and product
runtime dependencies. It does not install Java, EPUBCheck or any accessibility tool.

Download an official [EPUBCheck distribution](https://www.w3.org/publishing/epubcheck/)
and retain its adjacent library directory. Supply its JAR and a new output directory:

```sh
python3 Build/Epub/validate_epub.py book.epub edited-book.epub \
  --epubcheck-jar /path/to/epubcheck/epubcheck.jar \
  --output /path/to/task-evidence/epubcheck
```

The runner records the tool version and JAR hash, captures each input, hashes the
captured bytes, and retains the validator's log and JSON report. `summary.json`
identifies each input and outcome. A failed validator, timeout or missing report
produces a nonzero exit. Existing output directories are rejected to prevent mixing
reports from different runs. Inputs are limited to 256 publications, 128 MiB each;
`--timeout` controls the per-process limit in seconds.

To also capture automated accessibility checks, supply `--ace /path/to/ace-puppeteer`
from an independently installed [DAISY Ace](https://daisy.github.io/ace/) distribution.
Its browser runtime must already be available. The runner records its version, log,
JSON report, exit code, and reported outcome for the same captured EPUB bytes. Missing,
malformed, failing, or timed-out evidence produces a nonzero exit. Without `--ace`,
automated accessibility is recorded as unchecked.
On interruption or timeout, the runner terminates Ace's process tree, including detached
browser groups on POSIX. Cleanup has a bounded grace period beyond the requested audit
timeout. Run the POSIX timeout contract with
`python3 -m unittest discover -s Build/Epub`; it exercises ordinary and detached children.

A passing result establishes only the checks performed by the recorded tool versions.
Review warnings and retain the versions with the evidence. The summary explicitly
leaves comprehensive accessibility assessment and reader presentation unchecked even
when Ace passes. Complete human accessibility review, then inspect the same publication
bytes in the chosen reading systems. Never infer those results from validator exit codes.

Use a task-owned output location. Retain compact reports and decisive publication
fixtures deliberately; remove superseded captured publications and downloaded tools
when they are no longer needed. Source fixtures need producer, version and license
provenance before entering the maintained corpus.

## Typography fixtures

The opt-in generator uses the current EPUB owner to create Basic, Prose, and
Technical publications from one manuscript containing tables, code, long links,
an accessible SVG, Arabic, and Japanese. It also creates `glossary.epub` with two
terms, repeated references and localized return links, and `bibliography.epub` with
two entries formatted and title-sorted by `OfficeIMO.Bibliography`, three citations,
and return links. The bibliography fixture checks that order, escaped title text,
and italic formatting survive EPUB reopening. Its CSL style is an illustrative
local style, not a claim of qualification for every publisher's citation style.
It has no third-party test dependency and
is outside the normal solution and shipped packages.

`index.epub` exercises nested terms, multiple locators and an intra-index
cross-reference. Its labels target sections and chapters rather than inferred
screen page numbers. `renamed-index.epub` moves both chapters and the navigation
document into new directories, exercising locator, cross-reference and TOC repair.
Run both artifacts through the independent validators to compare their conformance.

`split-chapter.epub` divides a sectioned chapter into consecutive reading positions,
with forward, return and incoming links repaired across the new resource boundary.
Use it to check TOC order and link destinations in independent readers as well as
with the automated validators.

```sh
dotnet run --project Build/Epub/Fixtures/EpubFixtureGenerator.csproj -- \
  /path/to/task-evidence/typography
```

Supply a new directory. The generator writes EPUB files, a manifest with SHA-256
hashes, and expanded content with standards-mode HTML previews. The previews keep
the generated content and profile stylesheet but add simulated reader CSS for
light, dark, and enlarged text. Serve this directory on loopback when inspecting
the preview paths listed in `manifest.json`; for example:

```sh
python3 -m http.server 8766 --bind 127.0.0.1 \
  --directory /path/to/task-evidence/typography
```

Check narrow and wide widths, wrapping, direction, text enlargement, and inherited
colors. Run the EPUB files through `validate_epub.py`, and use the same captured
bytes for native reader checks. Browser previews do not reproduce EPUB pagination,
reader preferences, font handling, or assistive-technology behavior. Retain their
results separately from native reading-system acceptance.
