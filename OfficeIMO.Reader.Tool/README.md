# OfficeIMO.Reader.Tool

`OfficeIMO.Reader.Tool` exposes OfficeIMO's local document readers as a .NET global tool. It reads files or standard input, converts bounded folders concurrently, emits Markdown or the stable v5 JSON envelope, and materializes embedded assets.

## Install

```powershell
dotnet tool install --global OfficeIMO.Reader.Tool
```

## Read a file or standard input

```powershell
officeimo-reader read policy.docx --format markdown --output policy.md
officeimo-reader read report.pdf --format json --output report.reader.json --assets report-assets
Get-Content notes.md -Raw | officeimo-reader read - --name notes.md
```

`--output -` or an omitted output path writes to standard output. `--name` gives piped bytes an extension so Reader can choose the intended handler; its default is `stdin.txt`.

## Convert a folder

```powershell
officeimo-reader folder ./documents `
    --output ./converted `
    --format json `
    --assets ./converted-assets `
    --concurrency 4 `
    --max-files 500
```

Folder output preserves relative source paths and adds `.md` or `.reader.json`. Traversal is deterministic, skips directory reparse points, and is bounded by `--max-files`; `--max-total-bytes` provides an optional aggregate input ceiling. Use `--no-recursive` for the top directory only.

## Inspect capabilities

```powershell
officeimo-reader capabilities
officeimo-reader capabilities --format json
```

## Exit codes

| Code | Meaning |
| ---: | --- |
| 0 | Success |
| 2 | Invalid command or arguments |
| 3 | Input file or directory not found |
| 4 | Unsupported input format |
| 5 | Document discovery or reading failed |
| 6 | Output or asset materialization failed |
| 130 | Cancelled |

## Dependency boundary

The tool uses `OfficeIMO.Reader.All` and the existing local adapter graph. It has no third-party command parser, process launcher, native binary, model, network client, or hosted provider. OCR is not configured because an OCR engine is an explicit host dependency.
