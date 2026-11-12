---
title: "OfficeIMO CLI Documentation"
description: "Task guides for the officeimo command-line tool: install and pin it, convert and batch documents, extract content, work with PDFs and invoices, script it in CI, and run its MCP server."
layout: docs
---

`officeimo` is the command-line tool from the `OfficeIMO.Tool` NuGet package. It runs the same OfficeIMO engines as the .NET libraries, without Microsoft Office, from a shell script, a CI job, a scheduled task or an MCP client. These guides are organised by task. Each command also prints its own contract with `officeimo <area> --help`.

```shell
dotnet tool install --global OfficeIMO.Tool
officeimo convert report.docx
officeimo read report.docx --format markdown
officeimo hidden-data remove proposal.docx
```

## Guides

| Guide | Covers |
|---|---|
| [Install and pin the tool](/docs/cli/install/) | Global install, a repository tool manifest, updates and version checks |
| [Convert documents](/docs/cli/convert/) | One file at a time: Office, OpenDocument, Apple, email, HTML, text and spreadsheet routes |
| [Batch conversion and saved jobs](/docs/cli/batch-and-jobs/) | Folders, checkpoints, reusable job files, run reports and structured-data conversion |
| [Extract and inspect](/docs/cli/extract-and-inspect/) | Markdown and JSON extraction, compact inspection, workbook schemas and the JSON contracts |
| [PDF workflows](/docs/cli/pdf/) | Page copies, splits, decryption, OCR, page images, assembly, comparison, printing and redaction |
| [Before sharing](/docs/cli/before-sharing/) | Hidden data, Content Credentials and AI labels, and provenance checks in CI |
| [Invoices](/docs/cli/invoices/) | Inspect, validate, edit, convert and render CII and UBL invoices |
| [Scripting and CI](/docs/cli/scripting-and-ci/) | Exit codes, `--force`, standard output versus files, path rules, secrets and limits |
| [MCP and agents](/docs/cli/mcp-and-agents/) | The STDIO MCP server, allowed roots and the compact `agent` commands |

## Command areas

| Area | What it does |
|---|---|
| `convert` | Converts one file; the input and output extensions select the route |
| `read`, `extract`, `reader` | Extract documents or folders as Markdown or JSON |
| `inspect`, `agent` | Return compact, bounded JSON for inspection, search and fetch |
| `tabular` | List worksheets, report column schemas and convert CSV, TSV and workbooks |
| `workflow` | Batch conversion, saved jobs, page images, assembly, comparison, print plans and printing |
| `pdf` | Separate PDF copies, OCR and source-bound redaction |
| `html` | Render HTML, MHTML and ZIP site bundles to PDF or page images |
| `hidden-data`, `provenance` | Clean documents before sharing and check them in CI |
| `invoice` | Electronic invoice workflows |
| `markup` | Parse, validate, preview and export OfficeIMO Markup |
| `mcp` | Start the MCP server over standard input and output |

The command line covers the workflows listed here. The [.NET libraries](/libraries/) and [PSWriteOffice](/docs/pswriteoffice/) expose more of each engine; Studio features are not automatically CLI commands.
