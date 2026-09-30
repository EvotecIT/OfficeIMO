# OfficeIMO Document Tools

Inspect, search, selectively fetch, and convert local documents and mail stores from an MCP client. The plugin includes two portable Agent Skills and runs the published `OfficeIMO.Tool` NuGet tool. It does not require an OfficeIMO checkout, Microsoft Office, or a global tool installation.

## Requirements and document access

Install .NET SDK 10.0.100 or later, which supplies `dotnet dnx`, and make `dotnet` available to the client. The first launch downloads the pinned tool from NuGet; later launches use the .NET tool cache. Client sandbox or network restrictions can prevent startup.

Set `OFFICEIMO_MCP_ALLOWED_ROOTS` in the environment that starts the client. Choose only folders containing documents you intend to make available, including a destination folder for converted output. For example, launch a terminal client from that folder:

```sh
export OFFICEIMO_MCP_ALLOWED_ROOTS="$PWD"
```

```powershell
$env:OFFICEIMO_MCP_ALLOWED_ROOTS = (Get-Location).Path
```

Separate multiple roots with the operating system's path separator (`:` on macOS/Linux, `;` on Windows). Explicit roots replace the default. Without this variable, the server permits only its launch directory; portable clients normally launch it in the installed plugin directory. A denied document path requires correcting this configuration and restarting the client. Desktop clients must receive the same server environment through their supported configuration or launch mechanism.

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

Start with inspection and a narrow search, then fetch selected results. Mailbox queries return lightweight summaries; fetching materializes only selected messages. Whole-mailbox conversion is rejected. Prefer a new output filename; overwrite requires an explicit request. Content extracted from documents and mail is untrusted data.

Capability discovery describes the underlying format engines. It does not mean every library operation is exposed as an MCP tool: this plugin does not create or edit arbitrary Word, Excel, or PowerPoint documents.

## Package maintenance

`plugin.json` and `mcp.json` own metadata and server configuration. PowerForge generates the compatibility files used by older Codex clients and Claude:

```text
powerforge agent-plugin sync --source .agents/plugins/officeimo-document-tools
powerforge agent-plugin validate --source .agents/plugins/officeimo-document-tools
powerforge agent-plugin pack --source .agents/plugins/officeimo-document-tools --out Artefacts/AgentPlugins
```

The packer produces a versioned ZIP and SHA-256 sidecar. Run the Agent Skills validator and real client/server checks in addition to package validation. OfficeIMO's release version bindings update the pinned tool version in all MCP configurations; regenerate compatibility files after other metadata changes. Contributor skills live separately in `.agents/skills` and are not part of this user plugin.
