# OfficeIMO Markup VS Code Prototype

This extension is the first editor shell for `.office.md` and `.omd` files.

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

## Development

```powershell
npm install
npm run compile
.\scripts\dev-install.ps1 -Insiders -Force
```

For packaged Insiders installation:

```powershell
.\scripts\install-insiders.ps1 -Force
```

Packaged installs include a bundled Release build of `OfficeIMO.Markup.Cli`, so preview/export works even when the opened workspace is not the OfficeIMO source tree. Development links fall back to the sibling `OfficeIMO.Markup.Cli` project. Set `officeimoMarkup.cliPath` to an executable, DLL, or `.csproj` when testing another CLI build.

Open a `.omd` or `.office.md` file, run `OfficeIMO Markup: Open Preview`, or use the editor/explorer right-click menu. Plain `.md` files keep the built-in VS Code Markdown preview menu; invoking the OfficeIMO preview command on plain Markdown delegates back to VS Code's normal Markdown preview unless the file contains OfficeIMO front matter or directives. The OfficeIMO preview refreshes after edits while the panel is open, and the preview header now includes direct actions for `Refresh`, `Validate`, `Generate Artifacts`, `Export and Open`, and `Open Output Folder` so the preview works as a lightweight authoring control surface instead of only a renderer. The preview header also shows where generated files will land based on the active output-directory setting. Inline chart data and semantic layout blocks such as textboxes, columns, cards, and Mermaid diagrams get a lightweight live preview. Presentation preview keeps slides stacked one per row and avoids injecting debug metadata or extra generated text onto the slide canvas. Blank slides keep their semantic title in the AST but do not draw an extra preview title over authored layout blocks, matching PowerPoint export behavior. Chart blocks show single-series or grouped multi-series bars and surface exporter-facing metadata such as source ranges, source kind, target cell, chart size, axis titles, number formats, legend position, data labels, label format, and gridlines. Validation diagnostics are mapped back to the offending markup source line when the CLI returns node source text. Mermaid preview rendering is bundled into the VSIX and can be disabled with `officeimoMarkup.renderMermaidInPreview`; when preview rendering is unavailable, the raw diagram source remains visible. Presentation-profile files can be exported with `OfficeIMO Markup: Export Office Document` or the explicit profile commands: `OfficeIMO Markup: Export PowerPoint`, `OfficeIMO Markup: Export Excel Workbook`, and `OfficeIMO Markup: Export Word Document`. Every explicit export now offers `Open` and `Reveal` actions after success, `OfficeIMO Markup: Export and Open Office Document` uses the markup profile to export the right Office file and launch it immediately, the `Generate C# File` / `Generate PowerShell File` commands save code output instead of only opening scratch editors, `Open Generated C#` / `Open Generated PowerShell` jump straight back into the expected generated code files, and `Generate Artifacts` writes all three practical outputs in one pass: generated C#, generated PowerShell, and the final Office file chosen from the markup profile. The artifact success notification now lets you jump directly into the generated code files as well as the Office output. Use `officeimoMarkup.outputDirectoryMode` to choose whether generated files default beside the markup file or inside a configurable subfolder such as `generated`.

Run `OfficeIMO Markup: Install Mermaid Renderer` when you want Mermaid diagrams rendered into PNG images during PowerPoint export. The command installs `@mermaid-js/mermaid-cli` into this VS Code profile's extension storage and saves the discovered `mmdc` path in `officeimoMarkup.mermaidCliPath`. You can still set `officeimoMarkup.mermaidCliPath` manually to another Mermaid CLI executable. Leave it empty to use the extension-local renderer when available, then fall back to `PATH` and `OFFICEIMO_MARKUP_MERMAID_CLI`; if no renderer is found, export keeps readable diagram text.
