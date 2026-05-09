# OfficeIMO Markup for VS Code

Author Office documents from a Markdown-inspired format, preview the structure while you write, and export directly to PowerPoint, Excel, or Word using OfficeIMO.

OfficeIMO Markup is useful when you want repeatable document generation without opening Office first: meeting decks, workbook reports, narrative documents, generated examples, or source-controlled Office content.

## What You Get

- Syntax highlighting and snippets for `.omd` and `.office.md` files.
- Live preview for presentation, document, workbook, layout, chart, card, column, textbox, and Mermaid blocks.
- Inline validation while editing.
- Commands to generate C# or PowerShell code from markup.
- Export commands for `.pptx`, `.xlsx`, and `.docx`.
- Self-contained bundled CLI builds for Windows, Linux, and macOS on x64 and arm64.
- Optional Mermaid CLI integration for rendering diagrams during PowerPoint export.

## Quick Start

1. Create a file named `deck.office.md` or `deck.omd`.
2. Add OfficeIMO front matter and content.
3. Run `OfficeIMO Markup: Open Preview`.
4. Run an export command such as `OfficeIMO Markup: Export PowerPoint`.

```markdown
---
profile: presentation
title: Quarterly Update
---

# Quarterly Update

::: slide
## Highlights

- Revenue grew 18 percent
- Support backlog dropped
- Next quarter focuses on automation
:::
```

## Commands

- `OfficeIMO Markup: Open Preview`
- `OfficeIMO Markup: Validate`
- `OfficeIMO Markup: Generate C#`
- `OfficeIMO Markup: Generate PowerShell`
- `OfficeIMO Markup: Generate C# File`
- `OfficeIMO Markup: Generate PowerShell File`
- `OfficeIMO Markup: Generate Artifacts`
- `OfficeIMO Markup: Export Office Document`
- `OfficeIMO Markup: Export and Open Office Document`
- `OfficeIMO Markup: Export PowerPoint`
- `OfficeIMO Markup: Export Excel Workbook`
- `OfficeIMO Markup: Export Word Document`
- `OfficeIMO Markup: Open Output Folder`
- `OfficeIMO Markup: Open Generated C#`
- `OfficeIMO Markup: Open Generated PowerShell`
- `OfficeIMO Markup: Install Mermaid Renderer`

## Supported Files

The extension activates for:

- `.omd`
- `.office.md`
- Markdown files that contain OfficeIMO front matter or directives

Plain `.md` files keep the built-in VS Code Markdown preview unless the document looks like OfficeIMO Markup.

## Requirements

For normal packaged installs, no separate .NET runtime is required on the bundled platforms. The VSIX includes self-contained CLI executables for `win-x64`, `win-arm64`, `linux-x64`, `linux-arm64`, `osx-x64`, and `osx-arm64`.

Advanced users can set `officeimoMarkup.cliPath` to a custom `OfficeIMO.Markup.Cli` executable, DLL, or `.csproj`. DLL and project paths require a local .NET SDK or runtime.

Mermaid preview rendering is bundled in the extension. PowerPoint export can render Mermaid diagrams to PNG images when Mermaid CLI is available. Run `OfficeIMO Markup: Install Mermaid Renderer` to install `@mermaid-js/mermaid-cli` into the current VS Code profile's extension storage.

## Settings

- `officeimoMarkup.defaultProfile` - fallback profile when a file does not include front matter.
- `officeimoMarkup.cliPath` - optional custom CLI path.
- `officeimoMarkup.outputDirectoryMode` - write generated files beside the source file or into a generated subfolder.
- `officeimoMarkup.outputSubfolderName` - subfolder name for generated outputs.
- `officeimoMarkup.previewAutoRefresh` - refresh preview automatically while editing.
- `officeimoMarkup.renderMermaidInPreview` - render Mermaid diagrams in the preview webview.
- `officeimoMarkup.renderMermaidOnExport` - render Mermaid diagrams during PowerPoint export.
- `officeimoMarkup.mermaidCliPath` - optional path to `mmdc`.

## Development

```powershell
npm install
npm run compile
.\scripts\dev-install.ps1 -Insiders -Force
```

Development links use the source tree and fall back to the sibling `OfficeIMO.Markup.Cli` project. Set `officeimoMarkup.cliPath` to an executable, DLL, or `.csproj` when testing another CLI build.

## Package

```powershell
npm run package
```

`npm run package` calls `scripts/package-vsix.cjs`, which:

- installs extension dependencies with `npm ci`
- publishes self-contained `OfficeIMO.Markup.Cli` builds for Windows, Linux, and macOS
- publishes a framework-dependent CLI fallback for unsupported runtimes with .NET installed
- replaces the bundled CLI runtime folders under `tools/OfficeIMO.Markup.Cli`
- compiles the extension JavaScript
- writes `dist/officeimo-markup-<version>.vsix`

Use `npm run package` instead of running raw `vsce package`; the raw VSCE command does not refresh the bundled CLI and can ship stale dependencies. The npm package command uses the cross-platform Node wrapper in `scripts/package-vsix.cjs`; `scripts/package-vsix.ps1` remains available for CI and PowerShell users.

For packaged Insiders installation:

```powershell
.\scripts\install-insiders.ps1 -Force
```

## Marketplace Publishing

Local publish requires a Visual Studio Marketplace personal access token in `VSCE_PAT`:

```powershell
$env:VSCE_PAT = '<token>'
npm run publish:marketplace
```

CI packaging and publishing are handled by `.github/workflows/vscode-extension.yml`.

- Pull requests and pushes touching the extension or its runtime dependencies package the VSIX and upload it as the `officeimo-markup-vsix` artifact.
- Manual `workflow_dispatch` can set `publish_marketplace=true` to publish to the Visual Studio Marketplace.
- `VSCE_PAT` must be configured as a repository or organization secret before marketplace publishing is enabled.
- `pre_release=true` packages and publishes the extension as a VS Code pre-release.

## Support

Report bugs and feature requests in the [OfficeIMO issue tracker](https://github.com/EvotecIT/OfficeIMO/issues). Include the OfficeIMO Markup file, the command you ran, and the generated output or error message when possible.
