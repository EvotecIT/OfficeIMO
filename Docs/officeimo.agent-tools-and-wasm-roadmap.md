# OfficeIMO Agent Tools and WASM Roadmap

Date: 2026-06-22

This roadmap turns the PdfItDown comparison, the PSWritePDF retirement work, and the OfficeIMO browser conversion proof into a concrete OfficeIMO direction.

## Decision

OfficeIMO should have two complementary surfaces:

- A public static browser conversion playground on OfficeIMO.com, hosted by GitHub Pages or the existing static site pipeline.
- A local developer/agent tool surface made of Codex skills, a repo-local plugin, and a future MCP server.

These should share the same OfficeIMO core conversion APIs. The browser app should not become the automation backend, and the MCP server should not be required for public static conversion.

## Why This Split

The browser path is good for private, local-in-browser conversion of user-selected files. It needs byte and stream APIs, Blazor WebAssembly publishing, friendly diagnostics, and strict dependency control.

The MCP path is good for agent workflows: inspecting repository fixtures, comparing outputs, running conversion checks, surfacing package/release state, and helping PSWriteOffice retirement work. It can use local files and developer tooling, but it should still call OfficeIMO core APIs rather than duplicating conversion logic.

## Current Evidence

`Docs/officeimo.blazor-wasm-conversion-proof.md` records a local Blazor WebAssembly proof:

- Word DOCX to PDF works for basic and empty fixtures.
- Excel XLSX to PDF works for the tested fixture.
- PowerPoint PPTX to PDF works for the tested fixture.
- A richer Word fixture still exposes a Unicode/font embedding gap for private-use bullet glyphs.

That evidence says GitHub Pages-style hosting is viable for selected conversions, but the public feature needs explicit browser API wrappers, memory/bundle validation, and font diagnostics before it is marketed as production-grade.

## Agent Assets Added

The repo now has a plugin scaffold at:

```text
.agents/plugins/officeimo-document-tools/
```

The plugin exposes four skills:

- `officeimo-conversion-operator`
- `officeimo-build-release`
- `officeimo-website-wasm`
- `pswriteoffice-retirement`

The plugin intentionally does not declare an MCP server yet. The MCP manifest should be added when an actual `officeimo mcp serve` or equivalent entrypoint exists.

## Proposed MCP Server

Target command:

```powershell
officeimo mcp serve
```

Initial tools:

- `list_supported_formats`
- `convert_document`
- `inspect_document`
- `run_conversion_fixture`
- `compare_pdf_outputs`
- `explain_conversion_failure`

Initial resources:

- `officeimo://formats`
- `officeimo://fixtures`
- `officeimo://conversion-matrix`
- `officeimo://release-state`

Implementation rules:

- Reuse OfficeIMO core libraries directly.
- Keep file-system access explicit and scoped.
- Return structured diagnostics and artifact paths.
- Avoid hidden Office, LibreOffice, or native process dependencies.
- Keep PSWriteOffice-specific behavior out of the MCP server unless it is explicitly a PowerShell UX check.

## OfficeIMO.com WASM Path

Static app mount:

```text
Website/static/apps/officeimo-converter/
```

Public content route:

```text
/playground/
```

Docs route:

```text
/docs/converters/browser-playground/
```

The first public implementation should support drag/drop DOCX, XLSX, and PPTX input, convert to PDF in the browser, and return a downloadable file. It should report unsupported constructs and known font gaps without uploading the document anywhere.

## Implementation Phases

### Phase 1 - Repo-Owned Direction

- Keep the plugin and skills in `.agents/plugins/officeimo-document-tools/`.
- Keep the static app mount contract in `Website/static/apps/officeimo-converter/`.
- Publish a website docs page that explains the browser boundary and supported direction.

### Phase 2 - Browser App

- Add a Blazor WebAssembly project under a source folder chosen for website apps or samples.
- Publish its output into `Website/static/apps/officeimo-converter/`.
- Add a drag/drop UI for local files.
- Call only OfficeIMO byte and stream APIs.
- Validate with Playwright against the published static output.

### Phase 3 - MCP Server

- Add an OfficeIMO CLI or tool host with `mcp serve`.
- Start read-only with format listing, fixture inspection, and conversion matrix resources.
- Add mutating or artifact-producing conversion tools only after path scoping and output directory behavior are explicit.
- Add the plugin `.mcp.json` only after the server command is present and validated.

### Phase 4 - PSWriteOffice Retirement

- Use the MCP server and plugin skills to compare old PSWritePDF scenarios against OfficeIMO/PSWriteOffice replacements.
- Move reusable behavior into OfficeIMO first.
- Keep PSWriteOffice cmdlets as thin, friendly wrappers.
- Update retirement docs and support matrices from validated behavior.

## Validation Matrix

| Area | Required evidence |
| --- | --- |
| Plugin | `validate_plugin.py .agents/plugins/officeimo-document-tools` |
| Skills | Valid frontmatter and plugin validation |
| Website content | JSON validity, route content present, and static manifest present |
| Browser conversion | Blazor WASM publish plus real browser checks |
| PDF fidelity | Fixture-based `%PDF` output and known-gap diagnostics |
| PSWriteOffice retirement | Source and package-mode smoke tests where package boundaries are involved |

## Open Gaps

- Unicode font embedding for richer Word to PDF output.
- Browser memory limits for large workbooks and presentations.
- Published app bundle size and startup budget.
- A stable CLI/MCP host name and installation path.
- A conversion support matrix shared by docs, MCP resources, and the browser UI.
