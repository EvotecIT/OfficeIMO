# OfficeIMO Agents

This folder contains repo-owned agent assets.

- `plugins/marketplace.json` is the local marketplace entry for Codex plugin installation.
- `plugins/officeimo-document-tools/` is the self-contained plugin bundle.
- The plugin contributes a compact local STDIO MCP server for document/mailbox inspection, search, selected fetch, conversion, and capability discovery.
- Plugin skills are the canonical reusable instructions for using those tools and for OfficeIMO conversion, PDF, WASM website, build, release, and PSWritePDF retirement work.

The MCP server is provided by the versioned `OfficeIMO.Tool` .NET tool package.
