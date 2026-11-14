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
| `inspect`, `agent *`, `pdf *`, `hidden-data`, `provenance`, `invoice`, `workflow compare`, `workflow optimize-images`, `workflow export-pages`, `workflow assemble`, `workflow print-plan`, `workflow printers`, `workflow print`, `workflow print-job` | One JSON result (text with `--format text` where offered) |
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
| `markup emit --output`, `markup export` | Replaced without `--force`; write to a new path to keep the previous file |

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

## Discover commands from the tool

`officeimo help --format json` describes every command area, subcommand, usage line, key option, input and output format and exit code of the installed version as one JSON document with the schema `officeimo.tool.commands.v1`. Use it to check in a script that a command exists before relying on it, or to generate your own reference:

```shell
officeimo help --format json | jq -r '.groups[].subcommands[].usage[]'
officeimo help --format json | jq '.exitCodes'
```

```powershell
$catalog = officeimo help --format json | ConvertFrom-Json
$catalog.groups | Where-Object id -eq 'pdf' | ForEach-Object { $_.subcommands.name }
```

## CI recipes

Each recipe installs one exact version, so a new release changes nothing until you move the pin. The examples use `3.6.1`; replace it with the version you have tested.

### GitHub Actions: convert a folder to PDF

This job converts everything under `docs` to PDF, keeps a checkpoint so a re-run of the same commit continues where an interrupted run stopped, and uploads the PDFs and the per-file report:

```yaml
name: documents
on:
  push:
  workflow_dispatch:

jobs:
  pdf:
    runs-on: ubuntu-latest
    env:
      OFFICEIMO_VERSION: 3.6.1
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: '10.0.x'
      - name: Install officeimo
        run: dotnet tool install OfficeIMO.Tool --version "$OFFICEIMO_VERSION" --tool-path .tools
      - name: Restore an interrupted run of this commit
        uses: actions/cache/restore@v4
        with:
          path: |
            out/pdf
            .officeimo/checkpoint
          key: officeimo-pdf-${{ github.sha }}
      - name: Convert to PDF
        run: |
          mkdir -p reports
          .tools/officeimo workflow batch --input-directory docs --output out/pdf \
            --checkpoint .officeimo/checkpoint --retry-failed --format ndjson > reports/batch.ndjson
      - name: Keep the checkpoint for a re-run
        if: failure() || cancelled()
        uses: actions/cache/save@v4
        with:
          path: |
            out/pdf
            .officeimo/checkpoint
          key: officeimo-pdf-${{ github.sha }}
      - uses: actions/upload-artifact@v4
        if: always()
        with:
          name: documents
          path: |
            out/pdf
            reports
```

### Azure DevOps: convert a folder to PDF

Azure Pipelines can retry the conversion step in the same job. The checkpoint makes each retry reuse the PDFs that were already verified, and `--retry-failed` converts the failed files again:

```yaml
trigger:
  - main

pool:
  vmImage: ubuntu-latest

variables:
  OFFICEIMO_VERSION: 3.6.1

steps:
  - task: UseDotNet@2
    inputs:
      packageType: sdk
      version: 10.0.x

  - script: dotnet tool install OfficeIMO.Tool --version $(OFFICEIMO_VERSION) --tool-path .tools
    displayName: Install officeimo

  - script: |
      mkdir -p $(Build.ArtifactStagingDirectory)/reports
      .tools/officeimo workflow batch --input-directory docs --output $(Build.ArtifactStagingDirectory)/pdf \
        --checkpoint $(Agent.TempDirectory)/officeimo-checkpoint --retry-failed --format ndjson \
        > $(Build.ArtifactStagingDirectory)/reports/batch.ndjson
    displayName: Convert to PDF
    retryCountOnTaskFailure: 2

  - task: PublishPipelineArtifact@1
    condition: succeededOrFailed()
    inputs:
      targetPath: $(Build.ArtifactStagingDirectory)
      artifact: documents
```

### How checkpoints behave

- The source folder, the output folder and the checkpoint folder must be separate. Batches with a checkpoint use the `fail` conflict policy.
- A rerun with the same settings reuses every output whose receipt still matches and converts the rest. The NDJSON summary reports `reused` next to `completed`.
- Restore the output folder and the checkpoint together. An output that went missing after it was recorded fails that file instead of being converted again silently.
- Changing the conversion settings needs a new output folder and checkpoint. Execution limits such as `--concurrency` can change between runs.
- Files that failed stay failed on a rerun unless you pass `--retry-failed`.

### Gate and publish with a tool manifest

If the repository already pins the tool in a [tool manifest](/docs/cli/install/#pin-it-to-a-repository), restore it instead of installing a version in the workflow:

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
      - run: mkdir -p artifacts
      - run: dotnet officeimo provenance check ./docs --format sarif > artifacts/provenance.sarif
      - run: dotnet officeimo workflow batch --input-directory ./docs --output ./artifacts/pdf --format ndjson > artifacts/batch.ndjson
      - run: dotnet officeimo hidden-data remove ./docs/proposal.docx --output ./artifacts/proposal.docx
      - uses: actions/upload-artifact@v4
        with:
          name: documents
          path: artifacts
```

`provenance check` exits with code `1` when it finds dangerous Unicode, which stops the job before anything is published.
