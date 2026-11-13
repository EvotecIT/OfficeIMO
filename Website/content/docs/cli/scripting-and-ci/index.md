---
title: "Scripting and CI with the OfficeIMO CLI"
description: "Exit codes, overwrite rules, stdout versus files, path safety, password variables and limits for officeimo in scripts and CI jobs."
layout: docs
---

`officeimo` is built to be called from scripts. Results go to standard output, diagnostics to standard error, failures to a documented exit code, and existing files are protected unless you say otherwise.

## Exit codes

| Code | Meaning |
|---:|---|
| `0` | Success |
| `1` | The requested validation completed and found errors, for example `invoice validate` or `provenance check` |
| `2` | Invalid command or option, including a conversion with no matching route |
| `3` | Input was not found |
| `4` | Input is unsupported, or an I/O operation failed |
| `5` | The document operation failed, including batches with failed items |
| `6` | Output failed, including a refused overwrite, or conversion finished with error-severity diagnostics |
| `130` | Cancelled |

Findings are not failures: `hidden-data inspect`, `provenance inspect` and `provenance audit` return `0` when they run successfully, whatever they find. Use `invoice validate` or `provenance check` when you want findings to fail the step.

```shell
officeimo convert report.docx report.pdf
case $? in
  0) echo "converted" ;;
  6) echo "report.pdf already exists or the output failed" ;;
  *) echo "conversion failed" ; exit 1 ;;
esac
```

```powershell
officeimo invoice validate invoice.xml | Out-File invoice-result.json
if ($LASTEXITCODE -eq 1) { throw 'The invoice has validation errors.' }
elseif ($LASTEXITCODE -ne 0) { throw "officeimo failed with exit code $LASTEXITCODE." }
```

## Standard output and standard error

| Command | Standard output |
|---|---|
| `read`, `extract`, `reader read` without `--output` | The extracted Markdown or JSON |
| `convert` | The output path |
| `inspect`, `agent *`, `pdf *`, `hidden-data`, `provenance`, `invoice`, `workflow compare`, `workflow optimize-images`, `workflow export-pages`, `workflow assemble`, `workflow print-plan`, `workflow printers`, `workflow print` | One JSON result (text with `--format text` where offered) |
| `workflow batch`, `workflow job run` | The batch report as JSON, or NDJSON with `--format ndjson` |
| `tabular sheets`, `tabular schema` | Tab-separated lines; schema types are `string`, `integer`, `number`, `boolean` or `datetime` |

Diagnostics, progress and warnings always go to standard error, so redirecting standard output to a file or a JSON parser stays safe:

```shell
officeimo inspect report.docx | jq '.tableCount'
officeimo workflow batch --input-directory in --output out --format ndjson 2> batch.log | jq -c 'select(.type == "item" and .item.status != "Completed")'
```

```powershell
$result = officeimo hidden-data inspect proposal.docx | ConvertFrom-Json
$result.findings | Where-Object outcome -eq 'Found' | Select-Object kind, description
```

JSON results carry a schema identifier, such as `officeimo.conversion.batch.v2`, `officeimo.hidden-data.result.v1`, `officeimo.workflow.assemble.v1` or `officeimo.document.read-result`. Check it before you rely on a field.

## Overwrite rules

Commands refuse to replace an existing file or folder, so rerunning a script cannot silently overwrite earlier output:

| Command | To replace existing output |
|---|---|
| `convert`, `read`, `extract`, `reader read`, `reader folder`, `agent convert`, `tabular convert`, `html`, `pdf`, `hidden-data remove`, `provenance remove`, `workflow export-pages`, `assemble`, `compare`, `optimize-images` | `--force` |
| `workflow batch` | `--conflict replace`, or `--conflict rename` to keep both |
| `invoice` | Not possible; choose a new path |
| `workflow job run --report`, `--save-job` | Not possible; choose a new path |

Outputs are staged and published atomically. A conversion that fails or is rejected leaves an existing destination unchanged. The MCP `officeimo_convert` tool takes the same permission as an `overwrite` argument.

## Path safety

- The input and the output must be different files. Links are resolved before comparing, so a symbolic link to the source is rejected too.
- Folder destinations cannot be the source folder or contain it, and Apple package and redaction outputs cannot land inside or replace their inputs.
- Relative paths resolve from the current directory; paths inside a saved job resolve from the job file's folder.
- Folder discovery for `provenance audit` and `check` skips symbolic links and `.git`, `bin`, `obj` and `node_modules`.

## Passwords and secrets

Commands that need passwords read them from environment variables whose names you pass, never from the command line, and keep values out of results and saved jobs:

| Command | Option |
|---|---|
| `pdf decrypt`, `pdf redact`, `workflow compare` | `--password-env <name>` |
| `pdf redact` re-encryption | `--output-password-env <name>` |
| `workflow compare` second file | `--comparison-password-env <name>` |
| `workflow batch`, saved jobs | `--source-password-env <name>`, `--pdf-password-env <name>` |
| `mcp serve` | `--pdf-password-env <name>` admits a variable for MCP tool calls |

In CI, map a secret to the variable for the step that needs it:

```yaml
- run: dotnet officeimo pdf decrypt protected.pdf --password-env PDF_OWNER_PASSWORD --output clear.pdf
  env:
    PDF_OWNER_PASSWORD: ${{ secrets.PDF_OWNER_PASSWORD }}
```

## Resource limits

Commands have bounded defaults, so an unexpected file fails fast instead of exhausting a build agent. Raise a limit only for trusted input:

| Area | Defaults | Options |
|---|---|---|
| `convert` | 64 MiB input, 256 MiB output, 10,000,000 characters per XML part | `--max-input-bytes`, `--max-output-bytes`, `--max-characters-in-part` |
| `pdf` | 100 pages, 256 MiB input, 512 MiB output, 25 million pixels per page | `--maximum-pages`, `--maximum-input-bytes`, `--maximum-output-bytes`, `--maximum-pixels-per-page` |
| `workflow batch` | 64 MiB input and 256 MiB output per file | `--maximum-files`, `--maximum-input-bytes`, `--maximum-output-bytes`, `--maximum-xml-characters-in-part` |
| `invoice` | 256 inputs, 64 MiB combined input and output, 16 MiB per XML file | `--max-items`, `--max-input-bytes`, `--max-output-bytes` |
| `provenance audit`, `check` | 256 files | `--max-items` up to 10,000 |
| `agent` | Bounded JSON | `--max-output-characters 512..64000` |

## A complete CI job

```yaml
name: documents
on: [push]
jobs:
  publish:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: '10.0.x'
      - run: dotnet tool restore
      - run: dotnet officeimo provenance check ./docs --format sarif > provenance.sarif
      - run: dotnet officeimo workflow batch --input-directory ./docs --output ./artifacts/pdf --format ndjson > batch.ndjson
      - run: dotnet officeimo hidden-data remove ./docs/proposal.docx --output ./artifacts/proposal.docx
      - uses: actions/upload-artifact@v4
        with:
          name: documents
          path: artifacts
```

The tool version comes from the committed tool manifest; see [Install and pin the tool](/docs/cli/install/).
