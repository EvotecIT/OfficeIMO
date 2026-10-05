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

A passing result establishes only the checks performed by that EPUBCheck version.
Review warnings and retain the exact version with the evidence. The summary explicitly
leaves accessibility assessment and reader presentation unchecked. Run Ace and a
human accessibility review separately, then inspect the same publication bytes in the
chosen reading systems. Never infer those results from the EPUBCheck exit code.

Use a task-owned output location. Retain compact reports and decisive publication
fixtures deliberately; remove superseded captured publications and downloaded tools
when they are no longer needed. Source fixtures need producer, version and license
provenance before entering the maintained corpus.
