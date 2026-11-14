---
title: "Convert Documents with the CLI"
description: "Convert one file with officeimo convert: Office, OpenDocument, Apple, email, HTML and text routes, data formats, and Markdown or JSON extraction."
layout: docs
---

`officeimo convert` converts one file. The input and output extensions select a route from the shared OfficeIMO conversion catalog, the same catalog used by the .NET `OfficeIMO.Workflows` package and the `officeimo_convert` MCP tool. For folders and repeatable jobs, use [batch conversion](/docs/cli/batch-and-jobs/).

```shell
officeimo convert report.docx report.pdf
officeimo convert report.docx                      # writes report.pdf next to the source
officeimo convert workbook.xlsx --output published/workbook.pdf
officeimo convert deck.pptx deck.pdf --force       # replace an existing deck.pdf
```

The destination can be positional or passed with `--output` (`-o`), but not both. Without a destination, the tool writes a sibling `.pdf`. An existing destination is refused unless you pass `--force`. On success the command prints the output path on standard output; diagnostics go to standard error.

## See the routes your version runs

```shell
officeimo convert --help
```

The help ends with every route the installed tool runs, for example `DOCX -> PDF  .docx .docm .dotx .dotm -> .pdf`. If no route matches the two extensions, the command stops with exit code `2` and the message names the extensions it could not match.

## Common routes

Office and OpenDocument to PDF:

```shell
officeimo convert report.docx report.pdf
officeimo convert archive.doc archive.pdf
officeimo convert budget.ods budget.pdf
officeimo convert slides.odp slides.pdf
officeimo convert diagram.vsdx diagram.pdf
officeimo convert brochure.pub brochure.pdf
```

Between Office and OpenDocument, and into Word:

```shell
officeimo convert report.docx report.odt
officeimo convert report.odt report.docx
officeimo convert workbook.xlsx workbook.ods
officeimo convert deck.pptx deck.odp
officeimo convert notes.md notes.docx
officeimo convert page.html page.docx
officeimo convert letter.rtf letter.docx
```

From PDF to editable formats:

```shell
officeimo convert contract.pdf contract.docx
officeimo convert statement.pdf statement.xlsx
officeimo convert slides.pdf slides.pptx
officeimo convert report.pdf report.html
```

Email, EPUB, XPS, DjVu and HEIC to PDF:

```shell
officeimo convert message.eml message.pdf
officeimo convert book.epub book.pdf
officeimo convert document.xps document.pdf
officeimo convert scan.djvu scan.pdf
officeimo convert photo.heic photo.pdf
```

## Structured data

When both files are data formats, `convert` runs the structured-data routes between JSON, JSON Lines, YAML, TOML, XML, CSV, TSV and XLSX:

```shell
officeimo convert settings.yaml settings.json
officeimo convert settings.json settings.toml
officeimo convert export.xml export.yaml
officeimo convert orders.jsonl orders.xlsx
officeimo convert people.csv people.json
```

The [batch data conversion table](/docs/cli/batch-and-jobs/#structured-data) lists every source and target pair; `officeimo convert --help` lists the same routes. Each conversion reports what the target format could not hold on standard error. TOML has no null, for example, so null values are omitted and reported:

```text
Warning null-dropped [JSON to TOML] [note]: TOML has no null, so null values were omitted.
```

Add `--require-no-loss` to turn any reported loss into a failure. Nothing is written and the command returns exit code `5`.

These options shape the data:

| Option | Effect |
|---|---|
| `--select <path>` | Takes the rows from a path below the root, such as `data.items` or `results[0].rows`, for API responses and other wrapped data |
| `--columns <path,...>` | Fixes the table columns and their order for CSV, TSV and XLSX output, for example `--columns name,address.city` |
| `--root-name <name>` | Names the root element of XML output when the data has no single named root; the default is `root` |
| `--sheet <name>`, `--range <A1:B2>` | Reads one worksheet, or one range of it, from XLSX input. An unknown sheet or a malformed range exits with code `2`; the error for an unknown sheet lists the sheets the workbook has |

```shell
officeimo convert response.json items.csv --select data.items --columns id,name,address.city
officeimo convert people.json people.xml --root-name people
officeimo convert budget.xlsx budget.json --sheet Summary --range A1:D20
```

### CSV and workbooks

CSV or TSV to XLSX, and XLSX to CSV, use the spreadsheet routes:

```shell
officeimo convert people.csv people.xlsx
officeimo convert people.tsv people.xlsx
officeimo convert budget.xlsx budget.csv
```

Delimited input becomes one worksheet with typed columns; CSV output takes the first worksheet's values. The data options above do not apply to these routes. To choose a sheet or a delimiter, or to write XLSB or XLS, use [`officeimo tabular convert`](#spreadsheets-and-delimited-files).

## Route-specific options

| Option | Applies to | Effect |
|---|---|---|
| `--require-no-loss` | Routes that report fidelity loss, such as data routes and Draw, Publisher and DjVu to PDF | Rejects the conversion before publishing when OfficeIMO reports source or output loss. Not available for DOC, DOCX, TXT, XLSX or PPTX to PDF. |
| `--allow-legacy-loss` | DOC input | DOC conversion blocks known legacy import loss by default; this flag accepts the reported reductions. |
| `--text-encoding <name>`, `--tab-size 1..32` | TXT to PDF | Text input detects a Unicode BOM and otherwise uses strict UTF-8; these set the encoding and tab width explicitly. |
| `--diagram-layers screen\|print` | ODG and FODG to PDF | Chooses visible layers. Print-visible layers are the default; source page sizes, including blank pages, are kept. |
| `--text-only` | Email to PDF | Writes the Reader text projection instead of the styled message with header block, body, inline images and attachment list. |

```shell
officeimo convert report.txt report.pdf --text-encoding utf-8 --tab-size 4
officeimo convert diagram.fodg diagram.pdf --diagram-layers screen
officeimo convert scan.djv scan.pdf --require-no-loss
officeimo convert message.msg archive.pdf --text-only
```

Publisher conversion supports Publisher 2002 and later; convert Publisher 97, 98 and 2000 files to a later Publisher format first.

## Limits

Conversion is bounded by default: 64 MiB of input, 10,000,000 characters per XML part and 256 MiB of PDF output. Raise them for larger trusted documents:

```shell
officeimo convert large.xlsx large.pdf --max-input-bytes 268435456 --max-output-bytes 1073741824 --max-characters-in-part 50000000
```

The source is opened read-only and checked for package bombs. Output is staged and published atomically, so a rejected or failed conversion leaves an existing destination unchanged.

## Markdown and JSON output

A `.md` or `.markdown` destination, and a `.json` destination for a source without a data route, runs the Reader pipeline instead of a conversion route. The result is a semantic projection of the content, not a fixed-layout rendering, and every format listed by `officeimo reader capabilities` is accepted:

```shell
officeimo convert workbook.xlsx workbook.md
officeimo convert report.docx report.json
officeimo convert handbook.pdf handbook.md --assets handbook-assets
```

`--assets <directory>` writes embedded images and other assets beside the Markdown. JSON output uses the `officeimo.document.read-result` schema described in [Extract and inspect](/docs/cli/extract-and-inspect/#reader-json). For a data file such as YAML, TOML, XML, CSV or XLSX, a `.json` destination converts the data instead; use `read` to get its Reader JSON:

```shell
officeimo read settings.yaml --format json --output settings.reader.json
```

## Apple Pages, Numbers and Keynote

`.pages`, `.numbers` and `.key` inputs convert to their Office counterparts or PDF. Both ZIP files and macOS directory packages are accepted:

```shell
officeimo convert document.pages document.docx
officeimo convert budget.numbers budget.xlsx
officeimo convert presentation.key presentation.pptx
officeimo convert document.pages document.pdf
```

These conversions write JSON evidence with fidelity diagnostics, coverage and a source snapshot SHA-256. By default they reject partial editable reconstruction and previews without known complete coverage. Accept those outcomes explicitly with `--iwork-mode auto|editable|visual`, `--allow-partial`, `--allow-incomplete-preview` and `--normalize-worksheet-names`. PDF output requires complete editable reconstruction. The destination must be outside the source package.

## HTML and web archives

`officeimo html` renders HTML, MHTML and ZIP site bundles with more control than `convert`:

```shell
officeimo html convert report.html --output report.pdf --profile print-paged
officeimo html convert archive.mhtml --output archive.pdf
officeimo html render report.html --encoder png --output report-pages.zip
officeimo html render page.html --profile screen-full-page --encoder png --output page.zip
```

Local and remote resource reads are disabled by default; data URIs and resources inside MHTML or ZIP input remain available. Use `--stylesheet`, `--base-uri` and the `--font-*` options to supply resources explicitly. [HTML to PDF and images](/docs/cli/html/) covers profiles, page selection, site bundles, fonts and limits.

## Spreadsheets and delimited files

`officeimo tabular` converts CSV, TSV and workbooks without loading an editable workbook model:

```shell
officeimo tabular convert input.csv output.xlsx
officeimo tabular convert workbook.xlsb output.tsv --sheet Data
officeimo tabular convert pipe-delimited.csv output.csv --delimiter '|' --output-delimiter ','
officeimo tabular convert legacy.xls modern.xlsx
```

Workbook input can be XLSX, XLSM, XLTX, XLTM, XLAM, XLSB or XLS; workbook output is XLSX, XLSB or XLS. Select a sheet with `--sheet <name>` or `--sheet-index <zero-based-index>`; a sheet that does not exist exits with code `2`, and the error lists the sheets the workbook has. `.tsv` input always uses a tab unless `--delimiter` overrides it, and `--output-delimiter` controls CSV or TSV output independently. `--no-header` treats the first row as data. Existing destinations require `--force`.

## Word image optimization

```shell
officeimo workflow optimize-images input.docx --analyze
officeimo workflow optimize-images input.docx --output optimized.docx --mode both --dpi 144 --quality 85
officeimo workflow optimize-images first.docx second.docx --output-directory optimized --output-format pdf --mode both
```

`--analyze` reports embedded images without writing a file. Each output is staged and reopened before publication. The JSON result (`officeimo.workflow.optimize-images.v1`) has one entry per input under `files`, with the output path, byte counts and image diagnostics; `--format text` prints one summary line per file instead. `--output-format docx|doc|pdf` chooses the copies' format for `--output-directory`; with `--output`, the file extension chooses it.

## Related

- [Conversion guides](/convert/guides/) for fidelity notes per route
- [Format compatibility](/compatibility/)
- [Batch conversion and saved jobs](/docs/cli/batch-and-jobs/)
