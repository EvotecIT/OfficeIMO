# OfficeIMO Markup for VS Code

Visual Studio Code extension for authoring `.office.md` and `.omd` files with OfficeIMO Markup. The extension provides syntax highlighting, snippets, validation, live preview, code generation, and Office export commands for presentation, workbook, and document profiles.

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

## Authoring

Open a `.omd` or `.office.md` file, then run `OfficeIMO Markup: Open Preview` or use the editor, explorer, or preview toolbar actions. Plain `.md` files keep the built-in VS Code Markdown preview unless the file contains OfficeIMO front matter or directives.

The preview refreshes while the panel is open and exposes direct actions for refresh, validation, artifact generation, export-and-open, and opening the output folder. Inline chart data, semantic layout blocks, textboxes, columns, cards, and Mermaid diagrams get a lightweight live preview. Presentation previews keep slides stacked one per row and avoid drawing extra generated metadata onto the slide canvas.

Generated outputs can be written beside the markup file or into a configured subfolder via `officeimoMarkup.outputDirectoryMode` and `officeimoMarkup.outputSubfolderName`.

## Mermaid

Mermaid preview rendering is bundled into the VSIX. PowerPoint export can render Mermaid diagrams to PNG images when Mermaid CLI is available.

Run `OfficeIMO Markup: Install Mermaid Renderer` to install `@mermaid-js/mermaid-cli` into the current VS Code profile's extension storage and save the discovered `mmdc` path in `officeimoMarkup.mermaidCliPath`. You can also set `officeimoMarkup.mermaidCliPath` manually. When no renderer is available, export keeps readable diagram text.

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

`npm run package` calls `scripts/package-vsix.ps1`, which:

- installs extension dependencies with `npm ci`
- publishes `OfficeIMO.Markup.Cli` for `net8.0`
- replaces the bundled CLI under `tools/OfficeIMO.Markup.Cli`
- compiles the extension JavaScript
- writes `dist/officeimo-markup-<version>.vsix`

Use this script instead of running raw `vsce package`; the raw VSCE command does not refresh the bundled CLI and can ship stale dependencies.

For packaged Insiders installation:

```powershell
.\scripts\install-insiders.ps1 -Force
```

Packaged installs include the bundled Release build of `OfficeIMO.Markup.Cli`, so preview and export work even when the opened workspace is not the OfficeIMO source tree.

## Marketplace Publishing

Local publish requires a Visual Studio Marketplace personal access token in `VSCE_PAT`:

```powershell
$env:VSCE_PAT = '<token>'
npm run publish:marketplace
```

CI packaging and publishing are handled by `.github/workflows/vscode-extension.yml`.

- Pull requests and pushes touching `OfficeIMO.Markup*` package the VSIX and upload it as the `officeimo-markup-vsix` artifact.
- Manual `workflow_dispatch` can set `publish_marketplace=true` to publish to the Visual Studio Marketplace.
- `VSCE_PAT` must be configured as a repository secret before marketplace publishing is enabled.
- `pre_release=true` packages and publishes the extension as a VS Code pre-release.
