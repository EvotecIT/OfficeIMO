# ChartForgeX document delivery example

This executable places the same prepared service chart and topology in Word, Excel, PowerPoint and PDF. It uses the shared paired Graphite palette, the adapter's document typography style, full accessible descriptions and destination-aware insertion. A repeated light chart follows the dark variant, exercising reuse without treating an artifact identifier as an image-content identifier.

The visual viewport is 600 × 360 logical pixels, prepared for a 450 × 270 point placement with `OfficeVisualDocumentStyle.Default`. Axis and legend text is 11.25 points at that size. Word and PDF constrain natural-size visuals to the available content width. Excel fits an authored cell range and PowerPoint fits a layout box, keeping chart proportions and reserving separate heading and caption space. The workbook uses an authored Letter print layout. Dark visuals retain their own filled canvas on white document pages; PowerPoint slide backgrounds use the selected palette.

Run from the repository root with the adapter's ChartForgeX source-project configuration or a local qualified 2.0.0 package feed:

```powershell
$outputRoot = if ($env:EVOTEC_SCRATCH_ROOT) { $env:EVOTEC_SCRATCH_ROOT } else { [System.IO.Path]::GetTempPath() }
if (-not (Test-Path -LiteralPath $outputRoot -PathType Container)) { throw 'The configured output volume is not available.' }
$outputDirectory = Join-Path $outputRoot 'officeimo-chartforgex-delivery'
dotnet run --project ./OfficeIMO.ChartForgeX.Examples/OfficeIMO.ChartForgeX.Examples.csproj -c Release -f net10.0 -- $outputDirectory
```

The output directory contains the source SVG/PNG images, a five-page DOCX/PDF, a five-sheet XLSX, a five-slide PPTX, and a two-page editable topology VSDX. The JSON reports describe placement conversion and native Visio fidelity. Managed PNG previews reopen the saved files and render their document layouts; SVG previews expose the saved editable Visio geometry.

Managed Word previews use OfficeIMO's estimated pagination. Managed Excel and PowerPoint previews use their respective layout engines. Saved-PDF previews render the serialized PDF page content. These are distinct from Microsoft Office application rendering and must be assessed separately when an application-specific fidelity claim is needed.

On Windows with PowerPoint Desktop installed, reuse that output directory and explicitly request its existing reference-render lane:

```powershell
dotnet run --project ./OfficeIMO.ChartForgeX.Examples/OfficeIMO.ChartForgeX.Examples.csproj -c Release -f net10.0 -- $outputDirectory --native-powerpoint
```

That opt-in produces `powerpoint-desktop/` and a separate result report. Unavailable or failed native rendering returns a nonzero exit code; it does not fall back to managed images and claim native proof. Ordinary generation and managed rendering do not launch Office applications.

The editable Visio specimen uses prepared topology bounds and connector routes. Its report retains styling, font, accessibility and other conversion limits. Flow and sequence diagrams use their existing native reflow policy; this example does not establish exact editable layout preservation for those families.
