using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfVisualSelectedComparisonTests {
    [Fact]
    public void IndependentlySelectedLongDocumentsRetainOriginalOrdinalPairsAndUnmatchedPages() {
        byte[] expected = Pages(150), actual = Pages(151);
        var report = PdfVisualComparer.Compare(expected, actual, options: new() {
            ExpectedPages = PdfPageSelector.Parse("149,2,149"), ActualPages = PdfPageSelector.Parse("150,2"), MaxPages = 3
        });
        Assert.Equal(150, report.ExpectedPageCount); Assert.Equal(151, report.ActualPageCount);
        Assert.Equal(new[] { 149, 2, 149 }, report.ExpectedPageNumbers); Assert.Equal(new[] { 150, 2 }, report.ActualPageNumbers);
        Assert.Equal(new[] { 149, 2 }, report.Pages.Select(page => page.PageNumber));
        Assert.Equal(new[] { 150, 2 }, report.Pages.Select(page => page.ActualPageNumber));
        Assert.Equal(new[] { 149 }, report.UnmatchedExpectedPageNumbers); Assert.Empty(report.UnmatchedActualPageNumbers);
        Assert.False(report.IsMatch); Assert.True(report.IsSelectedScope);
        Assert.Equal(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(expected)).ToLowerInvariant(), report.ExpectedSha256);
        string html = report.ToHtmlGallery("<Review & report>", 2_000_000);
        Assert.Contains("&lt;Review &amp; report&gt;", html); Assert.Contains("#pair-0", html);
        Assert.Contains("Expected 149 / actual 150", html); Assert.Contains("150 total pages; selected: 149,2,149", html);
        Assert.Contains("Expected page 149 has no selected actual partner", html);
        Assert.Contains("does not detect semantic changes or moved pages", html);
        Assert.Equal(6, html.Split("data:image/png;base64,").Length - 1);
    }

    [Fact]
    public void EqualSelectedPagesDoNotClaimWholeDocumentsWereCompared() {
        var report = PdfVisualComparer.Compare(Pages(120), Pages(121), options: new() {
            ExpectedPages = PdfPageSelector.Parse("120"), ActualPages = PdfPageSelector.Parse("121"), MaxPages = 1
        });
        Assert.True(report.IsMatch); Assert.Single(report.Pages); Assert.Empty(report.StructuralDifferences);
        Assert.Equal(121, Assert.Single(report.Pages).ActualPageNumber);
        var reverse = PdfVisualComparer.Compare(Pages(3), Pages(4), options: new() {
            ExpectedPages = PdfPageSelector.Parse("1"), ActualPages = PdfPageSelector.Parse("2,last"), MaxPages = 2
        });
        Assert.Equal(new[] { 4 }, reverse.UnmatchedActualPageNumbers);
    }

    [Fact]
    public void ScopeLimitsCancellationAndBoundedGalleryRemainEnforced() {
        byte[] bytes = Pages(4);
        Assert.Throws<InvalidOperationException>(() => PdfVisualComparer.Compare(bytes, bytes, options: new() {
            ExpectedPages = PdfPageSelector.Parse("1,1,1"), ActualPages = PdfPageSelector.Parse("2"), MaxPages = 2
        }));
        Assert.Throws<ArgumentException>(() => PdfVisualComparer.Compare(bytes, bytes, PdfPageSelection.From(1), new() {
            ActualPages = PdfPageSelector.Parse("2")
        }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfVisualComparer.Compare(bytes, bytes, options: new() {
            ActualPages = PdfPageSelector.Parse("8")
        }));
        var cancelled = new CancellationToken(true);
        Assert.Throws<OperationCanceledException>(() => PdfVisualComparer.Compare(bytes, bytes, cancelled, options: new() {
            ExpectedPages = PdfPageSelector.Parse("1"), ActualPages = PdfPageSelector.Parse("2")
        }));
        var report = PdfVisualComparer.Compare(bytes, bytes);
        Assert.Throws<InvalidOperationException>(() => report.ToHtmlGallery(null, 256));
        Assert.Throws<OperationCanceledException>(() => report.ToHtmlGallery(null, 100_000, cancelled));
        Assert.Throws<PdfReadLimitException>(() => PdfVisualComparer.Compare(bytes, bytes, options: new() { MaxTotalPixels = 100 }));
    }

    private static byte[] Pages(int count) => PdfDocument.Create(document => {
        for (int index = 0; index < count; index++) document.Page(page => page.Size(150, 180));
    }).ToBytes();
}
