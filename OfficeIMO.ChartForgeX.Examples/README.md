# ChartForgeX document delivery example

This executable places the same prepared service chart and topology in Word, Excel, PowerPoint and PDF. It uses the shared paired Graphite palette with document-sized typography, full accessible descriptions and explicit placement dimensions. A repeated light chart follows the dark variant, exercising reuse without treating an artifact identifier as an image-content identifier.

The visual viewport is 600 × 360 logical pixels. Placements use at most 450 points of width and also respect the actual page or slide content width. Axis and legend text is 15 logical pixels, or 11.25 points at the full placement size. The workbook uses an authored Letter print layout; its picture, heading and full caption occupy separate rows. Dark visuals retain their own filled canvas on white document pages; PowerPoint slide backgrounds use the selected palette.

Run from the repository root with the adapter's ChartForgeX source-project configuration or a local qualified 2.0.0 package feed:

```powershell
dotnet run --project ./OfficeIMO.ChartForgeX.Examples/OfficeIMO.ChartForgeX.Examples.csproj -c Release -f net10.0 -- <scratch-root>/officeimo-chartforgex-delivery
```

The output directory contains the source SVG/PNG images, a five-page DOCX/PDF, a five-sheet XLSX, a five-slide PPTX, and a two-page editable topology VSDX. The JSON reports describe placement conversion and native Visio fidelity. Managed PNG previews reopen the saved files and render their document layouts; SVG previews expose the saved editable Visio geometry.

Managed Word previews use OfficeIMO's estimated pagination. Managed Excel and PowerPoint previews use their respective layout engines. Saved-PDF previews render the serialized PDF page content. These are distinct from Microsoft Office application rendering and must be assessed separately when an application-specific fidelity claim is needed.

On Windows with PowerPoint Desktop installed, explicitly request its existing reference-render lane:

```powershell
dotnet run --project ./OfficeIMO.ChartForgeX.Examples/OfficeIMO.ChartForgeX.Examples.csproj -c Release -f net10.0 -- <scratch-root>/officeimo-chartforgex-delivery --native-powerpoint
```

That opt-in produces `powerpoint-desktop/` and a separate result report. Unavailable or failed native rendering returns a nonzero exit code; it does not fall back to managed images and claim native proof. Ordinary generation and managed rendering do not launch Office applications.

The editable Visio specimen uses prepared topology bounds and connector routes. Its report retains styling, font, accessibility and other conversion limits. Flow and sequence diagrams use their existing native reflow policy; this example does not establish exact editable layout preservation for those families.
