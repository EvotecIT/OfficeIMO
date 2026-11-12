---
title: "MCP Server and Agent Commands"
description: "Run officeimo as a local MCP server with explicit allowed folders, and use its agent commands to inspect, search and convert documents."
layout: docs
---

The same `officeimo` executable is a local MCP server. Agents inspect, search, fetch and convert documents and mail through it, with every path limited to folders you allow. The `agent` commands give scripts the same bounded operations directly.

To add OfficeIMO to Claude, Codex, GitHub Copilot, Cursor or another client with a ready-made configuration, start at [OfficeIMO for AI agents](/agents/). This page covers the server and commands themselves.

## Start the server

```shell
export OFFICEIMO_MCP_ALLOWED_ROOTS="$HOME/Documents/contracts:$HOME/Documents/exports"
officeimo mcp serve --stdio
```

```powershell
$env:OFFICEIMO_MCP_ALLOWED_ROOTS = "$HOME\Documents\contracts;$HOME\Documents\exports"
officeimo mcp serve --stdio
```

`OFFICEIMO_MCP_ALLOWED_ROOTS` lists the folders the server may read and write, separated with `:` on macOS and Linux or `;` on Windows. Include a destination folder for converted output. The server refuses to start, with exit code `2`, when the variable is unset or empty, and it does not add the launch directory unless you list it.

With the tool installed globally, a client configuration can call `officeimo` directly:

```json
{
  "mcpServers": {
    "officeimo": {
      "type": "stdio",
      "command": "officeimo",
      "args": ["mcp", "serve", "--stdio"],
      "env": { "OFFICEIMO_MCP_ALLOWED_ROOTS": "/absolute/path/to/documents" }
    }
  }
}
```

Without an installation, clients launch a pinned package through `dotnet dnx`; the [agents page](/agents/) has those configurations, and the [OfficeIMO Document Tools plugin](https://github.com/EvotecIT/OfficeIMO/tree/master/.agents/plugins/officeimo-document-tools#readme) bundles the launcher with document and mailbox skills.

## Tools

| Tool | Does |
|---|---|
| `officeimo_inspect`, `officeimo_inspect_email` | Compact inspection of a document or mail file |
| `officeimo_search`, `officeimo_search_email` | Bounded search in a document or mailbox, with continuation |
| `officeimo_fetch` | Fetch selected results by id |
| `officeimo_convert` | Markdown or Reader JSON, Apple packages and every catalog route the CLI runs |
| `officeimo_capabilities` | Supported formats and operations |
| `officeimo_inspect_hidden_data`, `officeimo_remove_hidden_data` | The [hidden data](/docs/cli/before-sharing/) report and cleanup |
| `officeimo_pdf` | Extract, decrypt, flatten, optimize or sanitize into a separate PDF |
| `officeimo_pdf_split`, `officeimo_pdf_assemble`, `officeimo_pdf_export_pages` | Split, assemble PDFs and images, and export page images |
| `officeimo_pdf_ocr_providers`, `officeimo_pdf_ocr` | Searchable copies through a provider registered at startup |
| `officeimo_pdf_print_plan` | Sheet layout without printing |
| `officeimo_pdf_compare` | Rendered comparison of selected pages, saved as HTML |

Results contain a short text summary and compact structured content; the server does not publish whole documents or mailboxes as resources. Writing tools need an explicit output path inside the allowed roots and `overwrite=true` to replace a file. Printing to a device stays a CLI command.

## Trusted startup options

OCR providers and password variables are set when the server starts, not by tool calls:

```shell
officeimo mcp serve --stdio --ocr-provider-assembly ./providers/OfficeIMO.Ocr.Tesseract.dll --ocr-option executable=/usr/local/bin/tesseract --ocr-option tessdata=/usr/local/share/tessdata
officeimo mcp serve --stdio --pdf-password-env PDF_OWNER_PASSWORD
```

Tool calls can choose a registered provider, language, confidence and pages, but cannot load assemblies or set executable paths. A PDF tool can read only password variables admitted with `--pdf-password-env`; none are admitted by default, and password values never appear in arguments or results.

## Agent commands

`officeimo agent` runs the same operations from a shell and prints one compact JSON object. When `OFFICEIMO_MCP_ALLOWED_ROOTS` is unset, these direct commands keep normal file access, because you run them explicitly.

```shell
officeimo agent inspect report.docx --max-output-characters 6000
officeimo agent search report.docx --query "renewal date" --take 5
officeimo agent fetch --source-id <sourceId> --id <result id> --path report.docx
officeimo agent convert report.docx --output report.md --format markdown
officeimo agent capabilities --extension .pst
```

Search first, then fetch only the results you need:

```powershell
$search = officeimo agent search .\report.docx --query "renewal date" --take 5 | ConvertFrom-Json
officeimo agent fetch --source-id $search.sourceId --id $search.results[0].id --path .\report.docx
```

Search and fetch return a cursor when more is available; pass it back with `--cursor`. `agent convert` writes the full result to a file and returns only a small summary with the output path, size and SHA-256.

### Mailboxes

For PST, OST, OLM, EMLX, Mbox, MBX and folders of messages, `search` works on lightweight message summaries and `fetch` opens only the selected message. Filter with `--subject`, `--sender`, `--folder-id`, `--since`, `--before`, `--has-attachments` and `--is-read`. One search examines at most 10,000 summaries and reports `itemsScanned` and `scanLimitReached`.

`search-email` searches message content and resumes from a checkpoint:

```powershell
$checkpoint = $null
do {
    $arguments = @('agent', 'search-email', './archive.pst', '--query', 'renewal date',
        '--fields', 'Subject,TextBody,HtmlBody', '--take', '5', '--max-items-scanned', '1000')
    if ($checkpoint) { $arguments += '--checkpoint', $checkpoint }
    $page = officeimo @arguments | ConvertFrom-Json
    $page.results | ForEach-Object { $_.id }
    $checkpoint = $page.nextCheckpoint
} while ($checkpoint)
```

Fields are `Subject`, `Sender`, `Recipients`, `TextBody`, `HtmlBody`, `RtfBody`, `AttachmentNames`, `Bodies` or `All`; attachment content is not searched. Follow `nextCheckpoint` even when a page returns no matches, and keep the query, fields and filters unchanged when resuming. `isComplete` tells you whether the whole store was enumerated.

## Treat content as data

Results label source text with `contentTrust: "untrusted"` and include a `contentSafety` summary. Email bodies are checked for hidden HTML text and instruction-like language. Agents should treat document and mail content as data, never as instructions, and the host should enforce its own tool and data permissions.

## Related

- [OfficeIMO for AI agents](/agents/)
- [Extract and inspect](/docs/cli/extract-and-inspect/)
- [Scripting and CI](/docs/cli/scripting-and-ci/)
