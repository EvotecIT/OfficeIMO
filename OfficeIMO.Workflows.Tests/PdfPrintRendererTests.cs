using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfPrintRendererTests {
    [Fact]
    public void PreparationRejectsSheetOutsideTheDeliveryWorkingSetBudget() {
        PdfDocument document = PdfDocument.Create(c => c.Page(p => p.Size(200, 300)));
        Assert.Throws<InvalidOperationException>(() => PdfPrintRenderer.Prepare(document,
            new PdfPrintPlanRequest { InputPath = "snapshot.pdf", PaperSize = PageSizes.A4 },
            new PdfPrintRenderOptions { Dpi = 600, MaximumPixelsPerImage = 40_000_000 }));
    }

    [Fact]
    public void PreparationRejectsUndecodableSourceInsteadOfPrintingAnEmptySheet() {
        PdfDocument document = PdfDocument.Create(c => c.Page(p => p.Size(PageSizes.A4).Content(content =>
            content.Item(item => item.Paragraph(text => text.Text("This source must not disappear"))))));
        Assert.Throws<NotSupportedException>(() => PdfPrintRenderer.Prepare(document,
            new PdfPrintPlanRequest { InputPath = "snapshot.pdf", PaperSize = new PageSize(200, 300) },
            new PdfPrintRenderOptions { Dpi = 600, MaximumPixelsPerImage = 40_000_000 }));
    }

    [Fact]
    public void HighResolutionSheetsDecodeAtTheirPreparedDimensions() {
        PdfDocument document = PdfDocument.Create(c => c.Page(p => p.Size(200, 300)));
        PdfPreparedPrintDocument prepared = PdfPrintRenderer.Prepare(document,
            new PdfPrintPlanRequest { InputPath = "snapshot.pdf", PaperSize = PageSizes.A3 },
            new PdfPrintRenderOptions { Dpi = 300, MaximumPixelsPerImage = 18_000_000 });
        PdfRenderedPrintSheet sheet = Assert.Single(prepared.Sheets);
        OfficeRasterImage decoded = sheet.Decode(CancellationToken.None);
        Assert.InRange((long)decoded.Width * decoded.Height, 16_000_001, 18_000_000);
        Assert.True(decoded.Height > decoded.Width);
        Assert.ThrowsAny<OperationCanceledException>(() => sheet.Decode(new CancellationToken(true)));
    }

    [Fact]
    public void PreparationUsesAuthenticatedSnapshotAndKeepsSelectionOrder() {
        var encryption = new PdfStandardEncryptionOptions("print-user") {
            OwnerPassword = "print-owner", AllowedPermissions = PdfStandardPermissions.Print
        };
        byte[] bytes = PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Paragraph(p => p.Text("Printable without extraction permission")).ToBytes();
        var document = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "print-user" });
        PdfPreparedPrintDocument prepared = PdfPrintRenderer.Prepare(document, new PdfPrintPlanRequest {
            InputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf"),
            PaperSize = PageSizes.A4, Pages = "1", Orientation = PdfPrintOrientation.Landscape
        });
        Assert.Single(prepared.Sheets);
        Assert.True(prepared.Sheets[0].Plan.PaperSize.Width > prepared.Sheets[0].Plan.PaperSize.Height);
        Assert.True(OfficeImageReader.TryValidateContent(prepared.Sheets[0].GetPng(), "sheet.png", out var image));
        Assert.True(image.Width > image.Height);
        Assert.Equal(bytes, document.ToBytes());
        Assert.Throws<PdfPermissionDeniedException>(() => document.Reader.Text());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void PreparationRejectsExceededResourceBudgets(int budget) {
        PdfDocument document = PdfDocument.Create(c => {
            c.Page(p => p.Size(200, 300)); c.Page(p => p.Size(200, 300));
        });
        var limits = new PdfPrintRenderOptions();
        if (budget == 0) limits.MaximumPages = 1;
        if (budget == 1) limits.MaximumPixelsPerImage = 1;
        if (budget == 2) limits.MaximumOutputBytes = 1;
        Action prepare = () => PdfPrintRenderer.Prepare(document, new PdfPrintPlanRequest { InputPath = "snapshot.pdf" }, limits);
        if (budget == 0) Assert.Throws<InvalidOperationException>(prepare);
        else Assert.Throws<PdfReadLimitException>(prepare);
    }

    [Fact]
    public void CancelledPreparationDoesNotRenderSheets() {
        PdfDocument document = PdfDocument.Create(c => c.Page(p => p.Size(200, 300)));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfPrintRenderer.Prepare(document,
            new PdfPrintPlanRequest { InputPath = "snapshot.pdf" }, cancellationToken: cancellation.Token));
    }
}
