---
title: "Position Text and Graphics on PDF Pages"
description: "Choose between flowing PDF text, positioned stamps, arbitrary page canvases, and imported page overlays with explicit coordinates."
layout: docs
---

A `PdfText` run belongs to normal document flow. It can change emphasis, color, font, baseline, and links, but it does not independently choose a starting coordinate. Fixed positioning is a page-level operation.

## Choose the positioning surface

- Use `PdfParagraph`, `PdfText`, and rich runs when content should reflow naturally.
- Use `Add-OfficePdfStamp -X -Y` for one positioned text or image stamp on an existing PDF.
- Use `Add-OfficePdfCanvas` for arbitrary text, rich text, images, shapes, drawings, or tables at page-aware coordinates.
- Use `Add-OfficePdfPageOverlay` to place an imported PDF page in a target rectangle, optionally as an underlay.

## Draw rich text at a coordinate

Canvas coordinates use PDF points from the visual top-left of the page. The callback runs once per selected page. `PdfCanvasText` accepts the same visual `TextRun`, hashtable, and object shapes as the flowing rich-text commands. Link targets are excluded because canvas text is painted at a fixed position; use flowing `PdfText` when the text must be clickable.

```powershell
Add-OfficePdfCanvas -Path '.\Report.pdf' -OutputPath '.\Positioned.pdf' -Content {
    PdfCanvasText -Run @(
        TextRun 'Owner: ' -Bold
        TextRun 'Platform' -Color '#0F766E'
        TextRun '  |  REVIEW COPY' -Italic
    ) -X 36 -Y 24 -FontSize 10
}
```

The active page supplies the default text box, from `X/Y` to the remaining page edges. Add `-Width` and `-Height` only when the text must stay inside an exact rectangle. No native canvas object or typed .NET array is needed.

The [positioned-canvas recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-PositionedCanvas.ps1) adds rich text at two fixed positions on every page of a two-page document.

This coordinate system differs from lower-level PDF editing and redaction APIs that expose native bottom-left PDF coordinates. Follow the coordinate contract documented by the command being used rather than mixing values between surfaces.
