# OfficeIMO.ChartForgeX.Markdown

Render Mermaid fences as static diagrams in OfficeIMO Markdown workflows. This optional adapter uses [ChartForgeX.Markup.Mermaid](https://github.com/EvotecIT/ChartForgeX) for parsing, layout, SVG and PNG rendering. Word, PDF and PowerPoint engines keep their own document and image-placement behavior.

## Word and PDF

Materialize a parsed Markdown document before exporting it:

```csharp
using OfficeIMO.ChartForgeX.Markdown;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word.Markdown;

var document = MermaidMarkdownAdapter.Materialize(
    MarkdownReader.Parse(markdown), diagnostic => Console.WriteLine(diagnostic.Message));

using var word = document.ToWordDocument(new MarkdownToWordOptions { FitImagesToContextWidth = true });
word.Save("report.docx");
document.SaveAsPdf("report.pdf");
```

The transform replaces Mermaid fences with embedded PNG images, including fences inside lists, quotes and other Markdown containers. The Word example uses the existing context-width image policy so wide diagrams fit the page, including list and quote indentation. It preserves captions and supplies alternative text from the rendered artifact. Other fence languages remain unchanged. Invalid or unsupported diagrams retain their source fence and report diagnostics through the supplied callback.

For a converter that accepts `MarkdownReaderOptions`, add `MermaidMarkdownAdapter.CreateTransform()` to `DocumentTransforms`. Each application renders its own document, so the same transform can be reused.

## Static HTML

```csharp
using OfficeIMO.ChartForgeX.Markdown;
using OfficeIMO.MarkdownRenderer;

var options = new MarkdownRendererOptions();
MermaidMarkdownAdapter.ConfigureHtml(options);
string body = MarkdownRenderer.RenderBodyHtml(markdown, options);
```

Each diagram is an SVG image embedded through a data URI. This isolates SVG identifiers when the same diagram appears more than once and keeps the output deterministic. Images fit their container, captions remain visible, and failed diagrams remain ordinary code fences. The adapter disables Mermaid's browser runtime; other renderer plugins retain their own settings.

## Word and PowerPoint markup

Register the same image transform on the OfficeIMO Markup parser. The Word and PowerPoint exporters accept the resulting embedded images:

```csharp
using OfficeIMO.ChartForgeX.Markdown;
using OfficeIMO.Markdown;
using OfficeIMO.Markup;
using OfficeIMO.Markup.PowerPoint;

var reader = MarkdownReaderOptions.CreateOfficeIMOProfile();
reader.DocumentTransforms.Add(MermaidMarkdownAdapter.CreateTransform());
var parsed = OfficeMarkupParser.Parse(markdown, new OfficeMarkupParserOptions {
    Profile = OfficeMarkupProfile.Presentation,
    MarkdownOptions = reader
});
using var presentation = parsed.Document.ToPowerPointPresentation();
presentation.Save("report.pptx");
```

The parser preserves transformed fences within list items, slide bodies and directive columns (`::column`, `::left` and `::right`). Text-only presentation layouts use the normal block renderer when a list contains images or other nested content.

For Word markup, select `OfficeMarkupProfile.Document` and import `OfficeIMO.Markup.Word`, then call `parsed.Document.ToWordDocument()`. Word markup fits unsized embedded images to the current section and honors explicit image dimensions. Both markup exporters expose `AllowDataUriImages` and `MaximumDataUriImageBytes` in their options.

The PowerPoint exporter embeds the generated PNG bytes in process, preserves their aspect ratio and alternative text, and includes image captions. PNG and JPEG data URI images are enabled by default, with a 16 MiB decoded-image limit. Set `MarkupToPowerPointOptions.AllowDataUriImages` or `MaximumDataUriImageBytes` to apply a stricter policy. Existing file-path access rules continue to apply to file images.

## Rendering limits

The supported grammar and static approximations belong to [ChartForgeX's Mermaid support matrix](https://github.com/EvotecIT/ChartForgeX/blob/main/docs/mermaid-support-matrix.md). Source diagnostics refer to the Markdown input being parsed; nested OfficeIMO markup fragments have their own source locations. Unsupported source settings are reported by the renderer.

Reference this adapter only in applications that need static Mermaid rendering. Core OfficeIMO converters do not require the Mermaid package.
