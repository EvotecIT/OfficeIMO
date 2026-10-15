using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfInspectionProjectionHtmlOptionsTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void OpenedHtmlExportPreservesSkippedProjectionThroughPageSelection(bool skipImages, bool usePageRanges) {
        PdfDocument pdf = PdfDocument.Load(PdfInspectionProjectionFixture.Build(skipImages ? 32 : 1),
            skipImages ? new PdfLoadOptions { Limits = new PdfReadLimits { MaxContentOperands = 20 } } : null);
        var options = new PdfToHtmlOptions {
            ReadOptions = new PdfReadOptions { IncludeImages = !skipImages, IncludeFormWidgets = skipImages,
                PageSelection = PdfPageSelection.From(1) },
            PageRanges = usePageRanges ? new[] { PdfPageRange.From(1, 1) } : null
        };
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => pdf.ToHtmlResult(options));
        Assert.Contains("unrequested projection", error.Message);
    }

    [Fact]
    public void OpenedHtmlExportPreservesBothSkippedProjectionsBeforeOptionalImageBudget() {
        PdfDocument pdf = PdfDocument.Load(PdfInspectionProjectionFixture.Build(32),
            new PdfLoadOptions { Limits = new PdfReadLimits { MaxContentOperands = 20 } });
        Assert.Throws<PdfReadLimitException>(() => pdf.ToHtmlResult());
        var options = new PdfToHtmlOptions {
            ReadOptions = new PdfReadOptions { IncludeImages = false, IncludeFormWidgets = false },
            PageRanges = new[] { PdfPageRange.From(1, 1) }
        };
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => pdf.ToHtmlResult(options));
        Assert.Contains("unrequested projection", error.Message);
    }

    [Fact]
    public void OpenedHtmlExportFullDefaultRetainsSelectedImagesAndWidgets() {
        PdfDocument pdf = PdfDocument.Load(PdfInspectionProjectionFixture.Build());
        PdfHtmlConversionResult result = pdf.ToHtmlResult(new PdfToHtmlOptions {
            PageRanges = new[] { PdfPageRange.From(1, 1) }
        });
        Assert.Contains("Native page text", result.Value);
        Assert.Equal(new[] { 1 }, result.Summary.PageNumbers);
        Assert.Equal(1, result.Summary.ImageCount);
        Assert.Equal(1, result.Summary.FormWidgetCount);
    }
}
