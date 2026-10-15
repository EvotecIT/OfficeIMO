# Agent plugin distribution

OfficeIMO Document Tools packages two user workflows and a local MCP server. `plugin.json` and `mcp.json` in `.agents/plugins/officeimo-document-tools` are the source of truth; PowerForge generates the Codex and Claude compatibility files. Contributor skills in `.agents/skills` guide work on this repository and are not part of the user plugin.

## Release contract

Use the repository's PowerForge release configuration. Its version bindings align `OfficeIMO.Tool`, the plugin, Registry metadata and the pinned launcher; compatibility files regenerate in that transaction. Every bundle change, including skills or installation instructions, needs a new product patch release. Preserve existing versioned archives and checksums.

The Agent Plugin Package workflow attaches the ZIP and SHA-256 sidecar to an `OfficeIMO-v...` release. Qualify the published NuGet tool and client installation separately; an attached ZIP establishes artifact availability.

Verify the following on the stable candidate:

- PowerForge `agent-plugin validate` and `pack` against `OfficeIMO.Tool.csproj`, plus an independent checksum and archive-content check.
- Both Agent Skills with the skill validator, including bundled resources.
- Installation from a fresh client profile outside an OfficeIMO checkout: both skills, MCP initialization and the advertised tools.
- Inspect, search, selected fetch and requested conversion against synthetic inputs, with a separate output.
- Startup without allowed roots fails; out-of-root access is denied; overwrite stays explicit.
- Codex and Claude installation, Cursor portable-plugin loading and the documented manual MCP configuration. Record unavailable clients as unqualified lanes.
- An existing installation receives a new plugin version through the client's update path.

Use synthetic documents for review examples. Skills select relevant workflows; they do not grant file access or additional mutation authority.

## Official MCP Registry

Publish the signed NuGet tool first and read it back from NuGet. Validate `OfficeIMO.Tool/server.json` against the pinned official schema, check version agreement, and verify that the actual packaged README contains exactly `<!-- mcp-name: io.github.EvotecIT/officeimo -->`.

An authorized publisher runs `mcp-publisher login github` and `mcp-publisher publish OfficeIMO.Tool/server.json`. An authorized CI lane can use the official GitHub OIDC flow. Verify the returned name, version and package through the [Registry API](https://registry.modelcontextprotocol.io/docs). The Registry catalogs servers; it does not install skills or create vendor-directory listings. Follow the [official quickstart](https://modelcontextprotocol.io/registry/quickstart).

## OpenAI directory

Local STDIO MCP plugins use OpenAI's [desktop-only submission route](https://developers.openai.com/plugins/deploy/desktop-only), which is limited to approved partners. Confirm the publishing organization's allowlist status with the OpenAI representative, complete identity verification, and follow the [submission requirements](https://developers.openai.com/plugins/deploy/submission).

Upload the qualified release ZIP through the documented local-MCP flow, resolve automated findings, submit for review, and publish the approved version from the portal. Verify the exact-name listing and a fresh desktop installation. Local execution limits which conversations can use the plugin.

The current bundle launches a pinned NuGet tool rather than embedding its runtime. Confirm that choice against the desktop-only review requirements, which ask for the server and dependencies in the package. Local installation does not establish directory acceptance. A hosted HTTP server would require a separate deployment and file-access design decision.

## Claude

The repository's `.claude-plugin/marketplace.json` provides the direct Claude Code install route. Validate the plugin and marketplace with the installed CLI and exercise a fresh session. Use strict validation where supported.

For wider discovery, follow [Anthropic's publication guide](https://code.claude.com/docs/en/plugins/publish) and submit at [claude.ai/directory/manage](https://claude.ai/directory/manage). Check the component support for Claude Code, Cowork and claude.ai. The directory portal and `claude-plugins-official` have different admission routes.

## Cursor

Cursor supports root `plugin.json`, `mcp.json` and `skills/` through Agent Plugins. Test a release copy under `~/.cursor/plugins/local/officeimo-document-tools`, reload Cursor, and inspect the loaded skills and server. This plugin uses no root-variable placeholders, avoiding Cursor's documented difference in variable expansion.

Submit the qualified Git source at [Cursor Marketplace publishing](https://cursor.com/marketplace/publish), following the [plugin documentation](https://cursor.com/docs/plugins). Verify the reviewed listing and a fresh marketplace installation after acceptance.

## Submission material

Prepare the release ZIP/checksum, immutable source revision, supported platforms and runtimes, installation instructions, support contact, publisher-owned privacy information and reproducible prompts with synthetic inputs and expected results. Describe the MCP tools actually exposed; broader library capabilities do not imply arbitrary document creation or editing through this plugin.

Make the local execution requirement, explicit roots, optional OCR deployment and file-writing behavior clear. Keep submission identifiers, feedback, approved versions and verified listing URLs in the release handoff. Update website directory status only after verifying the listing; keep open discovery work in the single product roadmap.
