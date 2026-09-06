using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSplitBoundsTests {
    [Theory]
    [InlineData(1, 3)]
    [InlineData(2, 2)]
    [InlineData(int.MaxValue - 1, 1)]
    [InlineData(int.MaxValue, 1)]
    public void SplitHandlesPartSizesLargerThanTheDocument(int pagesPerPart, int expectedParts) {
        var document = PdfDocument.Create(compose => {
            compose.Page(page => page.Size(200, 300));
            compose.Page(page => page.Size(210, 300));
            compose.Page(page => page.Size(220, 300));
        });
        var parts = document.Pages.Split(pagesPerPart);
        Assert.Equal(expectedParts, parts.Count);
        Assert.Equal(new[] { 200D, 210D, 220D }, parts.SelectMany(part => part.Inspect().Pages).Select(page => page.Width));
        var result = document.Pages.SplitResult(pagesPerPart);
        Assert.True(result.Succeeded);
        Assert.Equal(expectedParts, result.Value!.Count);
    }
}
