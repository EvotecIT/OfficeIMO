# OfficeIMO Markup for VS Code

OfficeIMO Markup for VS Code helps author Office documents from a Markdown-inspired `.omd` or `.office.md` format. It previews structure while you write and exports through the OfficeIMO Markup CLI to PowerPoint, Excel, or Word.

## What it provides

- Syntax highlighting and snippets for `.omd` and `.office.md`.
- Live preview for presentation, document, workbook, layout, chart, card, column, textbox, and Mermaid blocks.
- Inline validation while editing.
- Commands to generate C# or PowerShell starter code from markup.
- Export commands for `.pptx`, `.xlsx`, and `.docx`.
- Self-contained bundled CLI builds for Windows, Linux, and macOS on x64 and arm64.
- Optional Mermaid CLI integration for PowerPoint diagram export.

## Quick start

1. Create `deck.office.md` or `deck.omd`.
2. Add OfficeIMO front matter and content.
3. Run `OfficeIMO Markup: Open Preview`.
4. Run an export command such as `OfficeIMO Markup: Export PowerPoint`.

```markdown
---
profile: presentation
title: Quarterly Update
---

# Quarterly Update

@slide {
  layout: title-and-content
}

- Revenue grew 18 percent
- Support backlog dropped
- Next quarter focuses on automation
```

## Commands

- `OfficeIMO Markup: Open Preview`
- `OfficeIMO Markup: Validate`
- `OfficeIMO Markup: Generate C#`
- `OfficeIMO Markup: Generate PowerShell`
- `OfficeIMO Markup: Generate Artifacts`
- `OfficeIMO Markup: Export Office Document`
- `OfficeIMO Markup: Export and Open Office Document`
- `OfficeIMO Markup: Export PowerPoint`
- `OfficeIMO Markup: Export Excel Workbook`
- `OfficeIMO Markup: Export Word Document`
- `OfficeIMO Markup: Open Output Folder`
- `OfficeIMO Markup: Install Mermaid Renderer`

## Requirements

Packaged installs include self-contained CLI executables for supported platforms. Advanced users can set `officeimoMarkup.cliPath` to a custom CLI executable, DLL, or `.csproj`; DLL and project paths require a local .NET SDK or runtime.

Mermaid preview rendering is bundled in the extension. PowerPoint export can render Mermaid diagrams to PNG when Mermaid CLI is available. Run `OfficeIMO Markup: Install Mermaid Renderer` to install `@mermaid-js/mermaid-cli` into the current VS Code profile storage.

## Development

```powershell
npm install
npm run compile
.\scripts\dev-install.ps1 -Insiders -Force
```

## Package

```powershell
npm run package
```

`npm run package` builds fresh `OfficeIMO.Markup.Cli` runtimes into the VSIX, compiles the extension JavaScript, and writes `dist/officeimo-markup-<version>.vsix`. The generated CLI binaries are removed after packaging and are not committed, so source and shipped tooling cannot drift apart. Use this command instead of raw `vsce package`.

For packaged Insiders installation:

```powershell
.\scripts\install-insiders.ps1 -Force
```

## Publishing

Local Marketplace publish requires `VSCE_PAT`:

```powershell
$env:VSCE_PAT = '<token>'
npm run publish:marketplace
```

CI packaging and publishing are handled by `.github/workflows/vscode-extension.yml`.

## Boundaries

- Extension UI, packaging, and VS Code commands belong here.
- Markup parsing and semantic behavior belongs in `OfficeIMO.Markup`.
- Command-line parse/export behavior belongs in `OfficeIMO.Markup.Cli`.
- Export fidelity belongs in the target exporter packages.

Report bugs and feature requests in the [OfficeIMO issue tracker](https://github.com/EvotecIT/OfficeIMO/issues).
