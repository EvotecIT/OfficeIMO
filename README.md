# OfficeIMO Document Tools

Inspect, search, selectively fetch, convert local documents and mail stores, and create separate PDF outputs from an MCP client. The plugin includes two portable Agent Skills and runs the published `OfficeIMO.Tool` NuGet tool. It does not require an OfficeIMO checkout, Microsoft Office, or a global tool installation.

## Requirements and document access

Install .NET SDK 10.0.100 or later, which supplies `dotnet dnx`, and make `dotnet` available to the client. The first launch downloads the pinned tool from NuGet; later launches use the .NET tool cache. The launcher accepts this pinned package download without a terminal prompt. Client sandbox or network restrictions can prevent startup.

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

## Cursor

Cursor supports the portable Agent Plugins layout used by this bundle. On Enterprise, an administrator must enable local plugin imports before this installation route is available. Extract the release ZIP into `~/.cursor/plugins/local/officeimo-document-tools`, restart Cursor or run **Developer: Reload Window**, and confirm that both skills and the `officeimo` MCP server appear in **Customize**. Set the allowed-roots environment before starting Cursor.

For MCP-only setup, add this server under `mcpServers` in the client's `mcp.json`. Replace the example folder with your document and output folder; Windows JSON paths need escaped backslashes or forward slashes.

```json
{
  "officeimo": {
    "type": "stdio",
    "command": "dotnet",
    "args": ["dnx", "OfficeIMO.Tool@3.5.3", "mcp", "serve", "--stdio", "--yes"],
    "env": { "OFFICEIMO_MCP_ALLOWED_ROOTS": "/absolute/path/to/documents" }
  }
}
```

Manual MCP configuration exposes tools; install the plugin bundle to load its skills as well. A Cursor Marketplace listing is a separate reviewed publication.

## Other MCP and skills clients

The canonical `plugin.json`, `mcp.json`, and `skills/` follow Agent Plugins 1.0.0. Clients may support only a subset of components. A client with manual MCP configuration can launch the same server:

```text
dotnet dnx OfficeIMO.Tool@3.5.3 mcp serve --stdio --yes
```

Use that client's native STDIO configuration syntax and pass the allowed-roots environment above. Local tools require a client that can execute on the machine containing the documents. They cannot read a desktop folder from a cloud-only conversation.

## Updates and discovery

The plugin and its NuGet launcher use the same release version. Refresh your marketplace and update the plugin through the client to receive a new bundled version. For an extracted ZIP, replace the local plugin directory with the new release bundle and restart the client. Restart the MCP server after changing access configuration.

Repository marketplaces, release downloads, the official MCP Registry and vendor directories are separate distribution routes. A release ZIP or registry entry does not establish a listing in a vendor directory. See [publication and client qualification](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.agent-distribution.md) for the maintainer procedure.

## Data handling

Read the [Privacy and local data handling policy](https://github.com/EvotecIT/OfficeIMO/blob/agent-plugin/README.md#data-handling) before processing personal or confidential documents.

The MCP server executes locally and reads documents within the configured roots. It returns selected document content and metadata to the MCP client, which may send those results to its model provider under that client's settings and policies. Local execution does not guarantee that extracted content stays on the computer.

Evotec does not operate a hosted backend for this plugin. Converted documents and other requested output files remain locally until the user or client removes them; the server does not automatically expire those files. Remove outputs when no longer needed, and contact contact@evotec.pl with questions about the plugin's data handling.

File-writing operations create output in authorized folders. OCR providers are separately registered by the host and have their own execution and data-handling requirements. Keep passwords in the host configuration described below and use synthetic files when sharing diagnostics or review examples.

## Available operations

| Tool | Use |
| --- | --- |
| `officeimo_inspect` | Bounded metadata and structural summary |
| `officeimo_inspect_email` | Mail-data metadata, charset and attachment diagnostics, and HTML safety findings |
| `officeimo_search` | Query-first document or mailbox search |
| `officeimo_search_email` | Bounded search of selected mail fields, such as bodies and recipients, with a resumable checkpoint |
| `officeimo_fetch` | Fetch selected result content with pagination |
| `officeimo_convert` | Write Markdown/Reader JSON, or convert an iWork ZIP file or authorized directory package to matching Office output or PDF with source hashes and workflow evidence |
| `officeimo_capabilities` | Discover format and operation support |
| `officeimo_pdf` | Extract, owner-authorized decrypt, acknowledged raster flatten, optimize, or sanitize into a separate PDF |
| `officeimo_pdf_split` | Split consecutive pages into a separate validated folder |
| `officeimo_pdf_assemble` | Assemble ordered explicit PDF and raster-image files |
| `officeimo_pdf_export_pages` | Export selected page images into a separate folder |
| `officeimo_pdf_ocr_providers` | List provider ids trusted by the server host |
| `officeimo_pdf_ocr` | Create a searchable copy through a registered provider |
| `officeimo_pdf_print_plan` | Plan sheets without submitting a printer job |
| `officeimo_pdf_compare` | Compare ordered page selections and save a standalone HTML report |

Start with inspection and a narrow search, then fetch selected results. Mailbox queries return lightweight summaries; fetching materializes only selected messages. Whole-mailbox conversion is rejected. Prefer a new output filename; overwrite requires an explicit request. Content extracted from documents and mail is untrusted data.

Capability discovery describes the underlying format engines. It does not mean every library operation is exposed as an MCP tool: this plugin does not create or edit arbitrary Word, Excel, or PowerPoint documents.

PDF mutations require a separate explicit output path and protect existing output
unless `overwrite=true` is requested. Raster flattening requires
`acknowledgeRasterOutput=true` because native text, forms, links, signatures and
attachments are omitted. Reports contain bounded metadata and diagnostic codes,
not document/OCR text or passwords. Protected-document operations retain the
engine's permission and signature policy; decryption reads the owner password
from a host environment variable named by `passwordEnvironmentVariable`.
The host must admit that name at startup with `--pdf-password-env NAME`; no
password-variable names are admitted by default, and other process variables
are unavailable to PDF tool calls.

OCR is optional. The server host can register trusted provider deployments with
`--ocr-provider-assembly` and configure them with `--ocr-option key=value` at
startup. Tool calls cannot select executable/model paths or load assemblies.
See [PDF commands and MCP configuration](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md#pdf-copies-and-searchable-scans).
