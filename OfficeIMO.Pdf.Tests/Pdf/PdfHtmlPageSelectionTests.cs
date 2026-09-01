using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfHtmlPageSelectionTests {
    [Fact]
    public void DocumentHtmlExport_AppliesPageRangesBeforeSemanticReading() {
        PdfDocument source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page 1"));
        for (int pageNumber = 2; pageNumber <= 1001; pageNumber++) {
            int capturedPageNumber = pageNumber;
            source.PageBreak().Paragraph(paragraph => paragraph.Text("Page " + capturedPageNumber));
        }
        PdfDocument document = PdfDocument.Load(source.ToBytes());
        var options = new PdfHtmlSaveOptions {
            ReadOptions = new PdfReadOptions {
                Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 1 }
            },
            PageRanges = new[] { PdfPageRange.From(1001, 1001) }
        };

        PdfHtmlConversionResult result = document.ToHtmlResult(options);

        Assert.Contains("Page 1001", result.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("Page 1<", result.Value, StringComparison.Ordinal);
        Assert.Equal(1001, result.Summary.SourcePageCount);
        Assert.Equal(1, result.Summary.RenderedPageCount);
        Assert.Equal(new[] { 1001 }, result.Summary.PageNumbers);
    }
}
