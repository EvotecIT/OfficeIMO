using OfficeIMO.Drawing;
using OfficeIMO.OneNote.Pdf;
using OfficeIMO.TestAssets;
using System.Threading;

namespace OfficeIMO.OneNote.Tests;

public sealed class OneNoteImageExportTextShapingTests {
    [Fact]
    public void OneNoteRasterExport_ReachesTheSharedTextShapingProvider() {
        var page = new OneNotePage {
            Title = "A",
            PageSize = OneNotePageSize.IndexCard
        };
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new OneNotePageRenderingOptions {
            DefaultFont = new OfficeFontInfo(ManagedTextShapingTestAssets.FamilyName, 12D),
            TextShapingProvider = provider,
            TextShapingLanguage = "ar-SA"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));

        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.Contains(provider.Requests, request =>
            request.Text == "A" &&
            request.Language == "ar-SA");
    }

    [Fact]
    public void OneNoteVisualPdfRasterStage_ReachesTheSharedTextShapingProvider() {
        var section = new OneNoteSection { Name = "Shaping" };
        section.Pages.Add(new OneNotePage {
            Title = "A",
            PageSize = OneNotePageSize.IndexCard
        });
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new OneNoteVisualPdfOptions {
            RasterScale = 0.25D,
            PageRendering = new OneNotePageRenderingOptions {
                DefaultFont = new OfficeFontInfo(ManagedTextShapingTestAssets.FamilyName, 12D),
                TextShapingProvider = provider,
                TextShapingLanguage = "ar-SA"
            }
        };
        options.PageRendering.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));

        OfficeIMO.Pdf.PdfDocumentConversionResult result =
            section.ToVisualPdfDocumentResult(options);

        Assert.NotEmpty(result.Value.ToBytes());
        Assert.Contains(provider.Requests, request =>
            request.Text == "A" &&
            request.Language == "ar-SA");
    }

    [Fact]
    public async Task OneNoteVisualPdfAsync_ObservesCancellationBeforeHostShaping() {
        var section = new OneNoteSection { Name = "Cancelled shaping" };
        section.Pages.Add(new OneNotePage {
            Title = "A",
            PageSize = OneNotePageSize.IndexCard
        });
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new OneNoteVisualPdfOptions {
            RasterScale = 0.25D,
            PageRendering = new OneNotePageRenderingOptions {
                DefaultFont = new OfficeFontInfo(ManagedTextShapingTestAssets.FamilyName, 12D),
                TextShapingProvider = provider
            }
        };
        options.PageRendering.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var stream = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => section.SaveAsVisualPdfAsync(stream, options, cancellation.Token));

        Assert.Empty(provider.Requests);
    }
}
