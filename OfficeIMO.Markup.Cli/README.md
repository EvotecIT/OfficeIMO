# OfficeIMO.Markup.Cli - markup command-line tooling

`OfficeIMO.Markup.Cli` is the command-line entry point for parsing, validating, emitting starter code from, and exporting OfficeIMO Markup files.

It is primarily a repository/development tool and a bundled runtime for the VS Code extension. It targets modern .NET and depends on the markup exporter packages.

## Commands

```powershell
dotnet run --project OfficeIMO.Markup.Cli -- parse OfficeIMO.Markup\Examples\presentation.omd --format json
dotnet run --project OfficeIMO.Markup.Cli -- validate OfficeIMO.Markup\Examples\presentation.omd
dotnet run --project OfficeIMO.Markup.Cli -- emit OfficeIMO.Markup\Examples\presentation.omd --target csharp --output presentation.cs
dotnet run --project OfficeIMO.Markup.Cli -- emit OfficeIMO.Markup\Examples\presentation.omd --target powershell --output presentation.ps1
dotnet run --project OfficeIMO.Markup.Cli -- export OfficeIMO.Markup\Examples\presentation.omd --target pptx --output presentation.pptx
dotnet run --project OfficeIMO.Markup.Cli -- export OfficeIMO.Markup\Examples\workbook.omd --target xlsx --output workbook.xlsx
dotnet run --project OfficeIMO.Markup.Cli -- export OfficeIMO.Markup\Examples\document.omd --target docx --output document.docx
```

## Export targets

- `pptx`: uses `OfficeIMO.Markup.PowerPoint`.
- `xlsx`: uses `OfficeIMO.Markup.Excel`.
- `docx`: uses `OfficeIMO.Markup.Word`.

## Useful switches

- `--mermaid-renderer <path-to-mmdc>`: render Mermaid diagrams during PowerPoint export.
- `--no-mermaid`: keep Mermaid blocks as text fallback.
- `--no-safe-preflight`: disable Excel save-time preflight.
- `--no-defined-name-repair`: disable Excel defined-name repair.
- `--no-openxml-validation`: disable Excel Open XML validation.

## Boundaries

- CLI orchestration belongs here.
- Parser and AST behavior belongs in `OfficeIMO.Markup`.
- Export behavior belongs in the target exporter packages.
- VS Code extension packaging belongs in `OfficeIMO.Markup.VSCode`.

## Targets and license

- Targets: `net8.0`, `net10.0`.
- License: MIT.
