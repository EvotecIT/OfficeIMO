OfficeIMO.Pdf — Zero‑Dependency PDF Builder
================================================

Goals
 - Zero dependencies; rely on PDF standard 14 fonts.
 - Fluent API mirroring OfficeIMO.Markdown patterns.
 - Basic page model, headings, paragraphs, page breaks.
 - Reasonable defaults (Letter, 1in margins, Courier 11pt for wrapping reliability).

Quick start

```csharp
using OfficeIMO.Pdf;

var pdf = PdfDoc.Create()
    .Meta(title: "Hello PDF", author: "OfficeIMO")
    .H1("OfficeIMO.Pdf")
    .Paragraph(p => p.Text("A tiny, zero‑dependency PDF builder."))
    .Paragraph(p => p.Text("This paragraph demonstrates automatic wrapping within page margins using the Courier standard font."))
    .Save("HelloWorld.pdf");
```

Current feature set
 - Pages: automatic paging with vertical flow.
 - Blocks: `H1`, `H2`, `H3`, `Paragraph(...)`, `PageBreak()`.
 - Fonts: standard 14 fonts; default `Courier` for predictable wrapping.
 - Metadata: Title, Author, Subject, Keywords.

Design notes
 - Text layout is simple vertical flow with word‑wrapping based on monospaced Courier metrics (600 units/em).
 - No compression; content streams are plain for readability.
 - Future: images (via ImageSharp opt‑in), shapes, tables, simple styles, and better font metrics.
