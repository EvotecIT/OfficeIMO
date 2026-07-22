using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfVisualComparerTests {
    [Fact]
    public void Compare_ReportsPixelDifferencesThresholdsIgnoredRegionsAndGallery() {
        byte[] expected = BuildPdf("Expected visual text");
        byte[] actual = BuildPdf("Changed visual text");

        PdfVisualComparisonReport exact = PdfDocument.Open(expected).CompareVisual(actual);
        PdfVisualPageComparison page = Assert.Single(exact.Pages);
        var ignoredOptions = new PdfVisualComparisonOptions();
        ignoredOptions.IgnoredRegions.Add(new PdfPixelRegion(0, 0, page.Width, page.Height));
        PdfVisualComparisonReport ignored = PdfVisualComparer.Compare(expected, actual, options: ignoredOptions);
        PdfVisualComparisonReport threshold = PdfVisualComparer.Compare(expected, actual, options: new PdfVisualComparisonOptions {
            AllowedDifferenceRatio = 1D
        });
        string gallery = exact.ToHtmlGallery("Review proof");

        Assert.False(exact.IsMatch);
        Assert.False(page.IsMatch);
        Assert.True(page.DifferentPixels > 0);
        Assert.True(page.DifferenceRatio > 0D);
        Assert.True(page.MaximumChannelDifference > 0);
        Assert.NotEmpty(page.DiffPng);
        Assert.True(ignored.IsMatch);
        Assert.True(threshold.IsMatch);
        Assert.Contains("Review proof", gallery, StringComparison.Ordinal);
        Assert.Equal(3, Count(gallery, "data:image/png;base64,"));
    }

    [Fact]
    public void Compare_ReportsPageCountAndDimensionStructureWithCenterAlignment() {
        byte[] expected = PdfDocument.Create(new PdfOptions { PageSize = new PageSize(300, 400) })
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .ToBytes();
        byte[] actual = PdfDocument.Create(new PdfOptions { PageSize = new PageSize(320, 420) })
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .ToBytes();

        PdfVisualComparisonReport report = PdfVisualComparer.Compare(expected, actual, options: new PdfVisualComparisonOptions {
            Alignment = PdfVisualPageAlignment.Center,
            AllowedDifferenceRatio = 1D
        });

        Assert.False(report.IsMatch);
        Assert.Contains(report.StructuralDifferences, difference => difference.StartsWith("PageCount:", StringComparison.Ordinal));
        Assert.Contains(report.StructuralDifferences, difference => difference.StartsWith("Page 1 dimensions:", StringComparison.Ordinal));
        PdfVisualPageComparison page = Assert.Single(report.Pages);
        Assert.Equal(320, page.Width);
        Assert.Equal(420, page.Height);
    }

    [Fact]
    public void Compare_EnforcesPagePixelOutputAndCancellationBudgets() {
        byte[] pdf = BuildPdf("Bounded visual comparison");

        PdfReadLimitException pixels = Assert.Throws<PdfReadLimitException>(() =>
            PdfVisualComparer.Compare(pdf, pdf, options: new PdfVisualComparisonOptions {
                MaxPixelsPerImage = 1
            }));
        Assert.Equal(PdfReadLimitKind.RenderPixels, pixels.Kind);

        PdfReadLimitException output = Assert.Throws<PdfReadLimitException>(() =>
            PdfVisualComparer.Compare(pdf, pdf, options: new PdfVisualComparisonOptions {
                MaxTotalOutputBytes = 1
            }));
        Assert.Equal(PdfReadLimitKind.RenderBytes, output.Kind);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            PdfVisualComparer.Compare(pdf, pdf, cancellationToken: cancellation.Token));
    }

    private static byte[] BuildPdf(string text) => PdfDocument.Create(new PdfOptions { PageSize = new PageSize(240, 180) })
        .Paragraph(paragraph => paragraph.Text(text))
        .ToBytes();

    private static int Count(string value, string token) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(token, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += token.Length;
        }

        return count;
    }
}
