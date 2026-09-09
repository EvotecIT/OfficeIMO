using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfPrintRendererTests {
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
