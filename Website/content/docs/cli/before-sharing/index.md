---
title: "Clean Documents Before Sharing"
description: "Remove hidden data and metadata with officeimo hidden-data, manage Content Credentials and AI labels, and check files in CI with SARIF."
layout: docs
---

Two command areas prepare files for people outside your team. `hidden-data` finds and removes what leaves with a file without being visible: author names, comments, tracked changes, hidden content, speaker notes, scripts and photo locations. `provenance` works on Content Credentials (C2PA), AI-source labels and suspicious Unicode, and can gate a repository in CI. Neither command changes its input.

## Hidden data

```shell
officeimo hidden-data inspect proposal.docx
officeimo hidden-data remove proposal.docx
officeimo hidden-data remove proposal.docx --output shared.docx --force --kinds default,provenance --tracked-changes accept
officeimo hidden-data remove poster.pdf --kinds default,provenance --keep-provenance c2pa
```

`inspect` changes nothing. `remove` writes `<name>.cleaned<ext>`, such as `proposal.cleaned.docx`, unless you pass `--output`. Inputs are Word, Excel, PowerPoint, Visio, OpenDocument, EPUB, PDF, JPEG, PNG, WebP, GIF and TIFF.

| Option | Effect |
|---|---|
| `--kinds` | `default`, `all` or category names such as `provenance`, `hidden-content`, `tracked-changes`, `image-metadata`, `attachments`, `active-content` and `bookmarks`, comma-separated |
| `--tracked-changes keep\|accept\|reject` | What to do with revisions |
| `--keep-provenance c2pa,external,ai-source` | Keep selected provenance carriers when `provenance` is removed |
| `--anonymous-author <name>` | Name that replaces authors on retained comments and revisions; the default is `Author` |
| `--remove-invalidated-signatures` | Allow removal that would invalidate an Office package signature |
| `--format json\|text` | JSON (the default) or readable text |

The default kinds remove scripts and other active content and keep tracked changes, hidden content, attachments, bookmarks and provenance. Signed PDFs are always refused. JSON results use the `officeimo.hidden-data.result.v1` schema, with one finding per item and its outcome:

```json
{"schema":"officeimo.hidden-data.result.v1","operation":"inspect","format":"Word","changed":false,
 "findings":[{"kind":"DocumentProperties","outcome":"Found","scope":"Document","location":"/docProps/core.xml","description":"Author","value":"..."}]}
```

[Remove hidden data and metadata](/docs/before-sharing/hidden-data/) explains each category and what is kept.

## Content Credentials and AI labels

```shell
officeimo provenance capabilities
officeimo provenance inspect report.docx
officeimo provenance inspect image.png --format text
officeimo provenance assess page.html
officeimo provenance remove report.docx --output report.cleaned.docx
officeimo provenance remove poster.png --output poster.shared.png --keep-c2pa
```

`inspect` reports carriers. `assess` also checks text integrity (turn it off with `--no-text-integrity`) and, when you supply a trusted `--c2patool` executable, verifies Content Credentials offline. `remove` keeps the input format and malformed carriers, and routes package changes through the owning OfficeIMO library. Keep carrier classes with `--keep-c2pa`, `--keep-external-c2pa` and `--keep-ai-source`, and skip embedded assets with `--no-embedded`. Findings are data: a successful inspection that finds credentials still returns exit code `0`.

Batch commands are sequential and bounded:

```shell
officeimo provenance batch inspect one.docx two.pdf --max-items 20
officeimo provenance batch remove one.docx two.pdf --output-directory cleaned
```

JSON results use `officeimo.provenance.capabilities.v2`, `officeimo.provenance.result.v2` and `officeimo.provenance.batch.v2`. Each report includes the SHA-256 of the exact bytes read and explicit check states such as `Completed`, `NotConfigured` or `Unsupported`, so a check that did not run is never reported as clean.

## Check working files in CI

`audit` and `check` assess files and folders without changing them. `check` fails with exit code `1` when its policy finds evidence:

```shell
officeimo provenance audit ./documents --include "*.html" --exclude "private/*" --format ndjson
officeimo provenance check ./documents --format sarif > provenance.sarif
officeimo provenance check ./site --fail-on carriers
```

`--fail-on dangerous-text` is the default and looks for potentially dangerous Unicode. `carriers` fails on provenance carriers and `any` on either. SARIF 2.1.0 output carries Unicode findings, carrier notes, input hashes, check coverage and execution failures, so code-scanning tools can show them. A carrier note is evidence, not an authenticity or AI verdict.

Discovery skips symbolic links and `.git`, `bin`, `obj` and `node_modules`. `--include` and `--exclude` take repeatable wildcards matched against paths relative to each root. The default is 256 assessed files, configurable with `--max-items` up to 10,000. Exceeding a bound, or finding no eligible files, fails the run instead of reporting a partial or empty success.

```yaml
- run: dotnet tool restore
- run: dotnet officeimo provenance check ./docs --format sarif > provenance.sarif
- uses: github/codeql-action/upload-sarif@v3
  if: always()
  with:
    sarif_file: provenance.sarif
```

`check` reads the working files; it does not inspect Git's index or fetch a website.

## Related

- [Prepare a document for sharing](/docs/before-sharing/)
- [Content Credentials and AI labels](/docs/before-sharing/provenance/)
- [PDF redaction](/docs/cli/pdf/#redaction)
