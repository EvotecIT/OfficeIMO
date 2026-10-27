using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReadScanOptimizationTests {
    [Theory]
    [InlineData("BT /F1 12 Tf 20 40 Td (Plain text) Tj ET", "Plain text")]
    [InlineData("% Do BI S f sh\nBT /F1 12 Tf 20 40 Td (Do BI S f sh) Tj ET", "Do BI S f sh")]
    public void TextOnlyPagesKeepTextWithoutInventingImagesOrVectors(string content, string expected) {
        byte[] source = CreatePage(content);
        PdfDocumentReadResult read = PdfDocument.Load(source).Read(new PdfLoadOptions { Profile = PdfReadProfile.Structured });
        PdfLogicalPage page = Assert.Single(read.Pages);
        Assert.Contains(expected, read.Text);
        Assert.Empty(page.Images);
        Assert.Equal(0, page.VectorPrimitiveCount);
    }

    [Fact]
    public void EmptyImageScanStillEnforcesContentOperationLimit() {
        byte[] source = CreatePage("q Q q Q q Q q Q");
        PdfReadPage page = Assert.Single(PdfReadDocument.Open(source,
            new PdfLoadOptions { Limits = new PdfReadLimits { MaxContentOperations = 3 } }).Pages);
        Assert.Throws<PdfReadLimitException>(() => page.GetImagePlacements());
    }

    [Fact]
    public void FormOnlyPaintRetainsNestedTextAndVectorEvidence() {
        byte[] source = CreatePage("/Form1 Do", "BT /F1 12 Tf 20 40 Td (Nested text) Tj ET 20 20 40 10 re S");
        PdfDocumentReadResult read = PdfDocument.Load(source).Read(new PdfLoadOptions { Profile = PdfReadProfile.Structured });
        PdfLogicalPage page = Assert.Single(read.Pages);
        Assert.Contains("Nested text", read.Text);
        Assert.True(page.VectorPrimitiveCount > 0);
    }

    [Theory]
    [InlineData("Encoding")]
    [InlineData("ToUnicode")]
    public void UnusedFontStreamsDoNotParticipateInEmptyPaintScans(string streamKey) {
        // An unused font may be inherited from a shared resource dictionary.
        // Its oversized declared decoded data must not affect an image scan.
        string source = PdfEncoding.Latin1GetString(CreatePage("q Q"))
            .Replace("/BaseFont /Helvetica", "/BaseFont /Helvetica /" + streamKey + " 7 0 R")
            .Replace("trailer\n", "7 0 obj\n<< /Length 1 /Filter /CCITTFaxDecode /DecodeParms << /Columns 2147483647 /Rows 2 >> >>\nstream\nx\nendstream\nendobj\ntrailer\n");
        PdfReadPage page = Assert.Single(PdfReadDocument.Open(PdfEncoding.Latin1GetBytes(source)).Pages);

        Assert.Empty(page.GetImagePlacements());
        Assert.Equal((0, 0), page.GetVisibleVisualPrimitiveCounts(Array.Empty<PdfStructuredTable>()));
    }

    [Fact]
    public void EmptyImageScanChargesFontAliasesToTheSharedWorkBudget() {
        PdfReadPage page = CreatePageWithFontAliases();
        Assert.Empty(page.GetImagePlacements()); // Prime the content summary.
        long consumed = 0;
        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
            page.GetImagePlacements(1, 1, units => {
                consumed += units;
                if (consumed > 1) throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingWork, 1, consumed);
            }, static () => { }));
        Assert.Equal(PdfReadLimitKind.UnderstandingWork, error.Kind);
        Assert.True(error.Actual > error.Limit);
    }

    [Fact]
    public void EmptyImageScanObservesCancellationBetweenFontAliases() {
        PdfReadPage page = CreatePageWithFontAliases();
        Assert.Empty(page.GetImagePlacements());
        using var cancellation = new CancellationTokenSource();
        long consumed = 0;
        Assert.Throws<OperationCanceledException>(() =>
            page.GetImagePlacements(1, 1, units => {
                consumed += units;
                cancellation.Cancel();
            }, cancellation.Token.ThrowIfCancellationRequested));
        Assert.Equal(1, consumed);
    }

    private static PdfReadPage CreatePageWithFontAliases() {
        string source = PdfEncoding.Latin1GetString(CreatePage("q Q"))
            .Replace("/F1 5 0 R", "/F1 5 0 R /F2 5 0 R /F3 5 0 R");
        return Assert.Single(PdfReadDocument.Open(PdfEncoding.Latin1GetBytes(source)).Pages);
    }

    private static byte[] CreatePage(string content, string? formContent = null) {
        string resources = "/Font << /F1 5 0 R >>" + (formContent == null ? "" : " /XObject << /Form1 6 0 R >>");
        string form = formContent == null ? "" :
            "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 200 200] /Resources << /Font << /F1 5 0 R >> >> /Length " +
            formContent.Length + " >>\nstream\n" + formContent + "\nendstream\nendobj\n";
        return PdfEncoding.Latin1GetBytes("%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << " + resources + " >> /Contents 4 0 R >>\nendobj\n" +
            "4 0 obj\n<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream\nendobj\n" +
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n" + form +
            "trailer\n<< /Root 1 0 R /Size 7 >>\n%%EOF");
    }
}
