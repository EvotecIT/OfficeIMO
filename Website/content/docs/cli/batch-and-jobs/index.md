---
title: "Batch Conversion and Saved Jobs"
description: "Convert folders with officeimo workflow batch: checkpoints, retries, conflict policies, JSON reports, saved jobs and data conversions."
layout: docs
---

`officeimo workflow batch` converts many files with one command. It selects a route for each file from the same catalog as [`officeimo convert`](/docs/cli/convert/), publishes each output independently and returns one JSON report. Files without an eligible route are counted as skipped.

## Convert a folder

```shell
officeimo workflow batch --input-directory ./Documents --output ./PDF
officeimo workflow batch report.pdf appendix.pdf --output ./HTML --target html
officeimo workflow batch --input-directory ./Inbox --output ./Word --target docx --source-extension .odt --source-extension .rtf
```

PDF is the default target. Relative paths resolve from the directory where the command runs. Folders are searched recursively unless you pass `--no-recursive`. The main options:

| Option | Effect |
|---|---|
| `--target <extension>` | Target format, such as `pdf`, `docx`, `xlsx`, `pptx`, `html` or a [data format](#structured-data) |
| `--route <id>` | Forces one catalog route, such as `json-yaml`; without `--target`, the route's own target is used |
| `--source-extension <ext>` | Limits selection to an extension; repeat it for several |
| `--conflict fail\|rename\|replace` | What to do when an output already exists; `fail` is the default |
| `--concurrency 1..32` | Parallel conversions |
| `--profile faithful\|lightweight\|printready\|textonly` | Output profile |
| `--maximum-files`, `--maximum-input-bytes`, `--maximum-output-bytes`, `--maximum-xml-characters-in-part` | Resource limits |
| `--require-no-loss`, `--allow-legacy-loss`, `--diagram-layers`, `--text-encoding`, `--tab-size`, `--pages` | Route options, applied only to the routes they concern |
| `--verbose` | Lists each skipped file on standard error instead of only counting them |

A batch fails when nothing was selected, which usually means a wrong folder or target. Pass `--allow-empty` when an empty run is expected. Item failures return exit code `5`; outputs that were already published stay in place.

Files that no route converts to the target are skipped. Standard error shows only a count, such as `Skipped 2 files: no conversion route to the target format, or excluded by --source-extension.`, and the report lists each skipped file. Add `--verbose` to print them on standard error as well.

## Checkpoints and retries

A checkpoint folder records each verified output, so a rerun reuses completed work instead of converting it again:

```shell
officeimo workflow batch --input-directory ./Documents --output ./PDF --checkpoint ./PDF-State
officeimo workflow batch --input-directory ./Documents --output ./PDF --checkpoint ./PDF-State --retry-failed
```

Checkpoint runs require `--conflict fail` and separate source, output and state folders. Before reusing an output, the run checks it. Changed inputs, settings or resources, altered or missing outputs and existing outputs without a verified receipt are reported for inspection rather than overwritten. Concurrency and resource limits can change between runs without discarding verified outputs.

## Reports on standard output

| Format | Output |
|---|---|
| Default | One aggregate JSON object with the `officeimo.conversion.batch.v2` schema, a `summary` and per-file `items` |
| `--format json` | The same report, versioned and bounded to 10,000 items or 16 MiB of item data, with an explicit truncation marker |
| `--format ndjson` | One JSON line per file as it completes (`"type":"item"`), then a final `"type":"summary"` line |

Human-readable progress, diagnostics and the skipped-file count go to standard error, so standard output stays parseable:

```shell
officeimo workflow batch --input-directory ./Documents --output ./PDF --format ndjson > batch.ndjson
```

If discovery or a file budget stops a run after some files completed, the report keeps those items and adds `summary.failureSummary`, and the command fails.

## Saved jobs

Save reviewed batch settings to a job file without converting anything, review the plan, then run the same job from any directory:

```shell
officeimo workflow batch --input-directory ./documents --output ./pdf --checkpoint ./conversion-state --save-job ./monthly.officeimojob.json
officeimo workflow job plan ./monthly.officeimojob.json
officeimo workflow job run ./monthly.officeimojob.json --format ndjson
officeimo workflow job run ./monthly.officeimojob.json --retry-failed
```

`job run` takes `--retry-failed`, `--allow-empty`, `--verbose`, `--format json|ndjson` and `--report`. Paths inside the job resolve against the job file's location. The file uses the `officeimo.conversion.job.v2` schema; unknown schema versions, unknown fields and unsupported settings are rejected. `job plan` reports selection, destinations and collisions without writing anything. `job run` revalidates the current inputs.

Job files never contain passwords, and `workflow batch` never takes them on the command line. Pass environment-variable names instead and set the values at run time:

```shell
officeimo workflow batch --input-directory ./protected --output ./pdf --source-password-env SOURCE_PASSWORD --pdf-password-env PDF_PASSWORD --save-job ./protected.officeimojob.json
```

PSWriteOffice runs the same job file with `Convert-OfficeDocument -JobPath ./monthly.officeimojob.json`.

### Keep a record of each run

`--report` writes a run report independently of the checkpoint state. Use a new file for each attempt, in an existing folder outside the input, output and checkpoint trees:

```shell
mkdir runs
officeimo workflow job run ./monthly.officeimojob.json --report ./runs/first.json
officeimo workflow job report ./runs/first.json
officeimo workflow job run ./monthly.officeimojob.json --retry-failed --report ./runs/retry.json
```

Each report records the job's fingerprint, timing, outcome and bounded per-file results. `job report` reads it without running the job or opening its files. A report from a process that exited mid-run shows as interrupted. If the report cannot be created, no conversion starts.

## Structured data

The structured-data routes convert between JSON, JSON Lines, YAML, TOML, XML, CSV, TSV and Excel worksheets. Convert one file with [`officeimo convert`](/docs/cli/convert/#structured-data), which also takes options for row selection, column order and the XML root name. Convert many files with a batch, passing the data format as `--target` or naming a route with `--route`:

```shell
officeimo convert settings.yaml settings.json
officeimo workflow batch --input-directory ./config --output ./json --target json --source-extension .toml
officeimo workflow batch export.xml --output ./converted --target yaml
officeimo workflow batch orders.jsonl --output ./converted --target xlsx
officeimo workflow batch people.csv --output ./converted --route csv-toml
```

| Source | Targets |
|---|---|
| JSON (`.json`) | JSON Lines, YAML, TOML, XML, CSV, TSV, XLSX |
| JSON Lines (`.jsonl`, `.ndjson`) | JSON, YAML, TOML, XML, CSV, TSV, XLSX |
| YAML (`.yaml`, `.yml`) | JSON, JSON Lines, TOML, XML, CSV, TSV, XLSX |
| TOML (`.toml`) | JSON, JSON Lines, YAML, XML, CSV, TSV, XLSX |
| XML (`.xml`) | JSON, JSON Lines, YAML, TOML, CSV, TSV, XLSX |
| CSV (`.csv`) | JSON, JSON Lines, YAML, TOML, XML |
| TSV (`.tsv`, `.tab`) | JSON, JSON Lines, YAML, TOML, XML |
| XLSX (`.xlsx`) | JSON, JSON Lines, YAML, TOML, XML |

Use `jsonl`, `yaml`, `toml`, `xml`, `csv`, `tsv`, `json` or `xlsx` as the `--target` value. CSV and TSV files with `--target xlsx`, and workbooks with `--target csv`, run the spreadsheet routes described in [Convert documents](/docs/cli/convert/#csv-and-workbooks). For a chosen sheet, delimiter, XLSB or XLS, use [`officeimo tabular convert`](/docs/cli/convert/#spreadsheets-and-delimited-files).

Each conversion reports what the target format could not hold. TOML has no null, for example, so null values are omitted and reported. Loss diagnostics appear on standard error and in the report item, with a `lossKind` and `hasLoss: true`:

```json
"report": {
  "target": "TOML",
  "itemsWritten": 1,
  "diagnostics": [
    { "code": "null-dropped", "severity": "Warning", "location": "note", "lossKind": "Omission",
      "message": "TOML has no null, so null values were omitted." }
  ],
  "hasLoss": true
}
```

Add `--require-no-loss` to turn any reported loss into a failure, in a batch or a single `convert`. The item is not published and the command returns exit code `5`:

```shell
officeimo workflow batch --input-directory ./config --output ./toml --target toml --require-no-loss
```

Data routes run in ordinary batches and saved jobs; checkpoint runs reject them.

## Related

- [Scripting and CI](/docs/cli/scripting-and-ci/) for exit codes and output handling
- [YAML](/docs/yaml/), [TOML](/docs/toml/) and [CSV](/docs/csv/) library guides
