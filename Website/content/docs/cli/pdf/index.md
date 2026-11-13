---
title: "PDF Workflows from the CLI"
description: "Extract, split, decrypt, flatten, optimize, sanitize, OCR, assemble, compare, print and redact PDFs with the officeimo command line."
layout: docs
---

Every PDF command writes a separate file or folder at a destination you name; the source is never changed. Existing destinations are refused unless you pass `--force`.

## Copies, splits and cleanup

```shell
officeimo pdf extract report.pdf --pages last,1-3 --output selected.pdf
officeimo pdf split report.pdf --pages-per-document 10 --output report-parts
officeimo pdf optimize report.pdf --output optimized.pdf
officeimo pdf sanitize report.pdf --output sanitized.pdf
officeimo pdf flatten report.pdf --output raster.pdf --acknowledge-raster-output --dpi 150
```

Page selections keep their order and repeats: `last,1-3`, `odd`, `even`, `all,!2`. Flattening replaces each page with its rendered appearance, so native text, forms, links, signatures and attachments are dropped; `--acknowledge-raster-output` is required to confirm that.

These commands print one JSON result on standard output with the status, artifact paths, sizes and diagnostic codes. Document text and passwords are never included:

```json
{"operation":"extract","status":"Completed","failureKind":"None","succeeded":true,
 "outputPath":"/work/selected.pdf","outputBytes":14295,"artifactCount":1,
 "artifacts":[{"path":"/work/selected.pdf","sizeBytes":14295}],
 "diagnostics":[{"code":"conversion-output-reopened","severity":"Information"}],"truncated":false}
```

Defaults allow 100 selected pages or split parts, 256 MiB of input, 512 MiB of output and 25 million pixels per rasterized page. Change them with `--maximum-pages`, `--maximum-input-bytes`, `--maximum-output-bytes` and `--maximum-pixels-per-page` within the ranges printed by `officeimo pdf --help`.

## Protected PDFs

Passwords are read from environment variables, never from command-line values:

```shell
export PDF_OWNER_PASSWORD='...'
officeimo pdf decrypt protected.pdf --password-env PDF_OWNER_PASSWORD --output clear.pdf
```

```powershell
$env:PDF_OWNER_PASSWORD = '...'
officeimo pdf decrypt protected.pdf --password-env PDF_OWNER_PASSWORD --output clear.pdf
```

Decryption needs the owner password. Other operations on protected documents keep the PDF engine's permission and signature rules.

## Searchable scans with OCR

The tool ships no OCR engine or model. Supply a trusted provider assembly and select its provider id. `OfficeIMO.Ocr.Tesseract` registers `tesseract-cli`, which calls a separately installed Tesseract executable:

```shell
officeimo pdf providers --ocr-provider-assembly ./providers/OfficeIMO.Ocr.Tesseract.dll
officeimo pdf ocr scan.pdf --output searchable.pdf --ocr-provider tesseract-cli --ocr-provider-assembly ./providers/OfficeIMO.Ocr.Tesseract.dll --ocr-language eng
```

OCR adds a text layer to the source pages. Configure the provider with repeated `--ocr-option key=value` entries and set a minimum recognition confidence with `--ocr-min-confidence 0..1`; `--pages` and `--dpi` limit the work. Recognition confidence is provider evidence, not a guarantee of correct text. The [Tesseract provider options](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Ocr.Tesseract/README.md) list the supported keys.

## Page images

```shell
officeimo workflow export-pages report.pdf --output report-pages --pages 1-3,last --image-format png
officeimo workflow export-pages report.pdf --output thumbnails --image-format webp --max-dimension 400
```

Image formats are `png`, `jpeg`, `webp`, `tiff` and `svg`, with `--dpi 36..600`. Every image is decoded and checked before the folder is published.

`export-pages`, like the other `workflow` commands except `batch`, writes one JSON result to standard output, so a script can read the published files without listing the folder:

```shell
officeimo workflow export-pages report.pdf --output report-pages | jq -r '.files[].path'
```

The result carries a `schema` such as `officeimo.workflow.export-pages.v1`, the `status`, and the `diagnostics`. Add `--format text` for a short readable summary.

## Assemble one PDF from mixed sources

```shell
officeimo workflow assemble cover.png report.docx appendices ./attachments.zip --output complete.pdf
```

Sources keep the order you give them. Folders and ZIP archives expand in deterministic path order, with entry, size and compression checks; `--no-recursive` stops folder recursion. Accepted inputs are PDF, DOCX, XLSX, PPTX, HTML, common raster images, folders and ZIP archives. The assembled PDF is reopened before it is published.

## Compare two revisions

```shell
officeimo workflow compare current.pdf revised.pdf --output review.html --expected-pages 120-124,last --actual-pages 121-125,last
```

The comparison renders the selected pages, up to 100 per side, and pairs them by position in the selection. The standalone HTML report shows page links, the selected scope, document totals, render limitations and SHA-256 values for both files. It compares rendered appearance; it does not detect semantic changes or moved pages. Use `--password-env` and `--comparison-password-env` for protected files.

The JSON result (`officeimo.workflow.compare.v1`) gives the report path in `outputPath`, its `outputSha256`, and the comparison verdict in `summary`:

```shell
officeimo workflow compare current.pdf revised.pdf --output review.html | jq -r '.summary'
```

## Print plans and printing

`print-plan` computes sheet layout without a printer driver, which makes it safe to run in CI:

```shell
officeimo workflow print-plan complete.pdf --paper A4 --pages-per-sheet 2 --scale fit
officeimo workflow print-plan report.pdf --pages-per-sheet 6 --page-subset odd --scale custom --custom-scale 40 --alignment bottom-right --margin 18 --color grayscale
```

The plan lists each sheet's paper size and where every page lands on it, in points:

```shell
officeimo workflow print-plan complete.pdf --pages-per-sheet 2 | jq '.sheets[] | {sheetNumber, pages: [.placements[].pageNumber]}'
```

Both `print-plan` and `print` accept 1, 2, 4, 6 or 9 pages per sheet, `--page-subset all|odd|even`, nine alignment positions, `--margin` or per-edge margins in points, and `--scale fit|actual|fill|custom` with `--custom-scale` from 1 to 1000 percent. Odd and even refer to the original page numbers within `--pages`.

To print, list the queues and submit the job:

```shell
officeimo workflow printers
officeimo workflow printers --capabilities "Office printer"
officeimo workflow print report.pdf --printer "Office printer" --pages 1-3 --pages-per-sheet 2 --copies 1 --duplex long
officeimo workflow print-job wait --printer "Office printer" --job Office_printer-42
```

`printers` returns the queues as JSON, with `name`, `isDefault` and `requiresOutputFile` for each. `--capabilities` returns the printer's media, default media, trays, duplex, color, collation, copies and resolutions, and `--paper-sources` returns only the `id` and `name` of each tray. `print` returns the queue receipt with `printer`, `jobId`, `sheetCount`, `copies` and any `ignoredAttributes`. `print-job status|wait|cancel` returns the job's `state`, `isTerminal` and `stateReasons`. Add `--format text` to any of them for plain lines:

```shell
officeimo workflow printers | jq -r '.printers[] | select(.isDefault) | .name'
```

The returned job identifier confirms that the queue accepted the job, not that it printed. After an interrupted submission, check the queue before retrying. Windows file printers need `--output-file` naming a new local file. On macOS and Linux, printing uses the CUPS command-line tools.

## Redaction

Redaction is a reviewed, three-step workflow: plan the matches, decide which candidates to remove, then apply and verify. A recipe names the rules:

```json
{
  "schema": "officeimo.pdf.redaction.recipe.v1",
  "rules": [ { "name": "account-number", "kind": "Literal", "value": "Account: 123-45-6789" } ]
}
```

```shell
officeimo pdf redact plan contract.pdf --recipe redaction.recipe.json --evidence contract.plan.json > plan.json
```

The plan output contains `sourceSha256`, `recipeSha256` and a `candidates` list with ids and geometry, never the matched text. Record the decision for every current candidate:

```json
{
  "schema": "officeimo.pdf.redaction.decisions.v1",
  "sourceSha256": "<from the plan>",
  "recipeSha256": "<from the plan>",
  "approvedCandidateIds": [ "<candidate id>" ],
  "rejectedCandidateIds": []
}
```

```shell
officeimo pdf redact apply contract.pdf --recipe redaction.recipe.json --decisions contract.decisions.json --output contract-redacted.pdf --evidence contract-redacted.evidence.json
officeimo pdf redact verify contract.pdf --recipe redaction.recipe.json --decisions contract.decisions.json --output contract-redacted.pdf --evidence contract-verify.evidence.json
```

`apply` re-plans the exact source and recipe, removes only approved candidates and publishes after verification passes. Evidence files record rule names, hashes, counts and geometry, not the redacted text. Outputs, evidence and manifests can never replace the PDF, recipe, decisions or batch request, even with `--force`. Unknown JSON members are rejected, so a misspelled field fails instead of falling back to a default.

`pdf redact batch --request <batch.json>` runs a recipe over a folder through the `officeimo.pdf.redaction.batch-request.v1` contract. Redaction can use an OCR provider for image text with the same `--ocr-provider-assembly`, `--ocr-provider`, `--ocr-language` and `--ocr-option` options; `pdf redact providers` lists what a provider assembly registers. For protected input, `--password-env` supplies the owner password and `--output-password-env` the new AES-256 password when the output is re-encrypted.

## Related

- [PDF guide](/docs/pdf/) and [PDF security](/docs/pdf/security/)
- [Browser PDF workflows](/pdf/)
- [Scripting and CI](/docs/cli/scripting-and-ci/)
