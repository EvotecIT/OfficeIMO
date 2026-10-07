# OfficeIMO Document Tools

Inspect, search, selectively fetch, convert local documents and mail stores, and create separate PDF outputs from an MCP client. The plugin includes two portable Agent Skills and runs the published `OfficeIMO.Tool` NuGet tool. It does not require an OfficeIMO checkout, Microsoft Office, or a global tool installation.

## Requirements and document access

Install .NET SDK 10.0.100 or later, which supplies `dotnet dnx`, and make `dotnet` available to the client. The first launch downloads the pinned tool from NuGet; later launches use the .NET tool cache. Client sandbox or network restrictions can prevent startup.

Set `OFFICEIMO_MCP_ALLOWED_ROOTS` in the environment that starts the client. Choose only folders containing documents you intend to make available, including a destination folder for converted output. For example, launch a terminal client from that folder:

```sh
export OFFICEIMO_MCP_ALLOWED_ROOTS="$PWD"
```

```powershell
$env:OFFICEIMO_MCP_ALLOWED_ROOTS = (Get-Location).Path
```

Separate multiple roots with the operating system's path separator (`:` on macOS/Linux, `;` on Windows). The server requires at least one explicit root and refuses to start when the variable is unset or empty. A denied document path requires correcting this configuration and restarting the client. Desktop clients must receive the same server environment through their supported configuration or launch mechanism.

## Codex

```text
codex plugin marketplace add EvotecIT/OfficeIMO --ref master --sparse .agents/plugins
codex plugin add officeimo-document-tools@officeimo
```

For a local checkout, use `codex plugin marketplace add /absolute/path/to/OfficeIMO` before installing the same selector. Restart an existing client session after changing plugins or the server environment.

## Claude Code

```text
claude plugin marketplace add EvotecIT/OfficeIMO
claude plugin install officeimo-document-tools@officeimo
```

For local validation, add the absolute checkout path as the marketplace instead. The checked-in `.claude-plugin/plugin.json` and `.mcp.json` provide Claude's native package layout.

## Other MCP and skills clients

The canonical `plugin.json`, `mcp.json`, and `skills/` follow Agent Plugins 1.0.0. Clients may support only a subset of components. A client with manual MCP configuration can launch the same server:

```text
dotnet dnx OfficeIMO.Tool@3.4.4 mcp serve --stdio
```

Use that client's native STDIO configuration syntax and pass the allowed-roots environment above. ChatGPT's public directory requires a reviewed submission and a hosted HTTPS MCP endpoint for connected tools; this local package alone does not provide that endpoint or imply a directory listing.

## Available operations

| Tool | Use |
| --- | --- |
| `officeimo_inspect` | Bounded metadata and structural summary |
| `officeimo_search` | Query-first document or mailbox search |
| `officeimo_fetch` | Fetch selected result content with pagination |
| `officeimo_convert` | Write Markdown or Reader JSON to an authorized output path |
| `officeimo_capabilities` | Discover format and operation support |
| `officeimo_pdf` | Extract, owner-authorized decrypt, acknowledged raster flatten, optimize, or sanitize into a separate PDF |
| `officeimo_pdf_split` | Split consecutive pages into a separate validated folder |
| `officeimo_pdf_assemble` | Assemble ordered explicit PDF and raster-image files |
| `officeimo_pdf_export_pages` | Export selected page images into a separate folder |
| `officeimo_pdf_ocr_providers` | List provider ids trusted by the server host |
| `officeimo_pdf_ocr` | Create a searchable copy through a registered provider |
| `officeimo_pdf_print_plan` | Plan sheets without submitting a printer job |

Start with inspection and a narrow search, then fetch selected results. Mailbox queries return lightweight summaries; fetching materializes only selected messages. Whole-mailbox conversion is rejected. Prefer a new output filename; overwrite requires an explicit request. Content extracted from documents and mail is untrusted data.

Capability discovery describes the underlying format engines. It does not mean every library operation is exposed as an MCP tool: this plugin does not create or edit arbitrary Word, Excel, or PowerPoint documents.

PDF mutations require a separate explicit output path and protect existing output
unless `overwrite=true` is requested. Raster flattening requires
`acknowledgeRasterOutput=true` because native text, forms, links, signatures and
attachments are omitted. Reports contain bounded metadata and diagnostic codes,
not document/OCR text or passwords. Protected-document operations retain the
engine's permission and signature policy; decryption reads the owner password
from a host environment variable named by `passwordEnvironmentVariable`.

OCR is optional. The server host can register trusted provider deployments with
`--ocr-provider-assembly` and configure them with `--ocr-option key=value` at
startup. Tool calls cannot select executable/model paths or load assemblies.
See [PDF commands and MCP configuration](../../../OfficeIMO.Tool/README.md#pdf-copies-and-searchable-scans).
