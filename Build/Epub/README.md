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
