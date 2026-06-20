using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using System;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class MarkdownSaveAsPdfVisualTests {
    [Fact]
    public void ToPdfDocument_MarkdownImageBlock_RendersDataUriImageAsStyledFigure() {
        string dataUri = CreateDataUriPng();
        MarkdownDoc document = MarkdownDoc
            .Create()
            .H1("Visual report")
            .Image(dataUri, "Operational badge", "Badge", width: 42, height: 28)
            .Caption("Figure 1. Operational badge");

        byte[] bytes = document.ToPdfDocument(CreateVisualOptions()).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("Figure 1. Operational badge", text, StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
    }

    [Fact]
    public void ToPdfDocument_MarkdownImageOnlyParagraph_RendersImageInsteadOfTextPlaceholder() {
        string dataUri = CreateDataUriPng();
        MarkdownDoc document = MarkdownDoc
            .Create()
            .Add(new ParagraphBlock(new InlineSequence().Image("Inline badge", dataUri, "Inline badge caption")));

        byte[] bytes = document.ToPdfDocument(CreateVisualOptions()).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(bytes), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.Contains("Inline badge caption", text, StringComparison.Ordinal);
        Assert.DoesNotContain("[Image:", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_MarkdownChartFence_RendersChartVisualInsteadOfJsonCodePanel() {
        const string markdown = """
# Quarterly report

```chart
{
  "type": "bar",
  "title": "Quarter revenue",
  "data": {
    "labels": ["Q1", "Q2", "Q3"],
    "datasets": [
      { "label": "Actual", "data": [10, 14, 19], "backgroundColor": "#2563EB" }
    ]
  },
  "width": 360,
  "height": 220
}
```
_Figure 2. Revenue chart_
""";

        var options = CreateVisualOptions();
        byte[] bytes = markdown.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "UnsupportedChartFence");
        Assert.Contains("Quarter revenue", text, StringComparison.Ordinal);
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Q1", text, StringComparison.Ordinal);
        Assert.Contains("Figure 2. Revenue chart", text, StringComparison.Ordinal);
        Assert.DoesNotContain("\"datasets\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_CustomReaderOptionsWithoutSemanticChartExtension_LeavesChartFenceAsCodePanel() {
        const string markdown = """
```chart
{
  "type": "bar",
  "data": {
    "labels": ["Q1"],
    "datasets": [
      { "label": "Actual", "data": [10] }
    ]
  }
}
```
""";

        var options = CreateVisualOptions();
        options.ReaderOptions = MarkdownReaderOptions.CreateOfficeIMOProfile();

        byte[] bytes = markdown.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "UnsupportedChartFence");
        Assert.Contains("\"datasets\"", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_InvalidMarkdownChartFence_WarnsAndFallsBackToSemanticPanel() {
        const string markdown = """
```chart
{ invalid json
```
""";

        var options = CreateVisualOptions();
        byte[] bytes = markdown.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains(options.Warnings, warning => warning.Code == "UnsupportedChartFence");
        Assert.Contains("{ invalid json", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ToPdfDocument_UnsupportedVisualFence_UsesSemanticFallbackPanelWithWarning() {
        const string markdown = """
```mermaid
graph TD
A[Markdown AST] --> B[OfficeIMO PDF]
```
_Figure 3. Flow fallback_
""";

        var options = CreateVisualOptions();
        byte[] bytes = markdown.ToPdfDocument(options).ToBytes();
        string text = PdfCore.PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains(options.Warnings, warning => warning.Code == "UnsupportedSemanticFence" && warning.Source == MarkdownSemanticKinds.Mermaid);
        Assert.Contains("mermaid", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("graphTD", text, StringComparison.Ordinal);
        Assert.Contains("Figure 3. Flow fallback", text, StringComparison.Ordinal);
    }

    private static MarkdownPdfSaveOptions CreateVisualOptions() => new MarkdownPdfSaveOptions {
        VisualTheme = MarkdownPdfVisualTheme.Report(),
        PdfOptions = new PdfCore.PdfOptions {
            CompressContentStreams = false,
            PageWidth = 420,
            PageHeight = 420,
            MarginLeft = 36,
            MarginRight = 36,
            MarginTop = 36,
            MarginBottom = 36
        }
    };

    private static string CreateDataUriPng() {
        string base64 = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 1));
        return "data:image/png;base64," + base64;
    }
}
