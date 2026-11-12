---
title: "OfficeIMO.Tool"
description: "Convert, extract, inspect and clean documents from your terminal with the officeimo CLI. Install the .NET tool from NuGet and use it in scripts, CI jobs, batch conversions and MCP clients."
meta.seo_title: "OfficeIMO.Tool: document conversion and inspection CLI"
layout: application
meta.application: "tool"
---

## One command for document work in scripts

`officeimo` runs the OfficeIMO engines from a shell, a CI job, a scheduled task or an AI agent, without Microsoft Office. Results go to standard output as JSON, diagnostics go to standard error, failures return documented exit codes, and existing files are protected unless you pass `--force`.

## Install

```shell
dotnet tool install --global OfficeIMO.Tool
officeimo --version
officeimo help
```

For scripts and CI, pin the version in the repository with a tool manifest and run it as `dotnet officeimo`. [Install and pin the tool](/docs/cli/install/) shows both.

## Convert a folder of documents to PDF

```shell
officeimo workflow batch --input-directory ./Documents --output ./PDF --checkpoint ./PDF-State
```

Each file is converted, reopened and published independently. Rerun the same command to reuse completed work, or add `--retry-failed`. Single files use `officeimo convert report.docx`. See [Convert documents](/docs/cli/convert/) and [Batch conversion and saved jobs](/docs/cli/batch-and-jobs/).

## Extract content for search or a model

```shell
officeimo read report.docx --format markdown
officeimo reader folder ./contracts --output ./contracts-json --format json
officeimo inspect deck.pptx
```

Markdown and versioned JSON come from the same Reader used by the .NET libraries. See [Extract and inspect](/docs/cli/extract-and-inspect/).

## Clean a document before sharing

```shell
officeimo hidden-data inspect proposal.docx
officeimo hidden-data remove proposal.docx --tracked-changes accept
officeimo provenance check ./docs --format sarif > provenance.sarif
```

`hidden-data` writes a cleaned copy without author names, comments, scripts and photo locations; `provenance check` fails a CI step when working files contain potentially dangerous Unicode, or provenance carriers if you ask for that. See [Before sharing](/docs/cli/before-sharing/).

## Prepare PDFs

```shell
officeimo pdf extract report.pdf --pages last,1-3 --output selected.pdf
officeimo workflow assemble cover.png report.docx ./attachments.zip --output complete.pdf
```

The PDF commands also split, decrypt, flatten, OCR, compare, print and redact. See [PDF workflows](/docs/cli/pdf/).

## Use it from an AI agent

```shell
officeimo mcp serve --stdio
```

The MCP server reads and writes only inside the folders named in `OFFICEIMO_MCP_ALLOWED_ROOTS`. See [MCP and agents](/docs/cli/mcp-and-agents/) or [OfficeIMO for AI agents](/agents/) for client setup.

## All guides

The [CLI documentation](/docs/cli/) covers every command area, including [invoices](/docs/cli/invoices/) and [scripting and CI](/docs/cli/scripting-and-ci/). Each command also prints its own reference with `officeimo <area> --help`.
