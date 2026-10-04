using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfTextEditorTests {
    [Theory]
    [InlineData("BT 50 650 Td (keep me) Tj ET\n")]
    [InlineData("0 0 40 40 re f\n")]
    public void MutationRejectsPromotedActualTextPaintWhoseStateFeedsLaterContent(string followingContent) {
        byte[] source = BuildRawTextPdf(
            "/Span << /ActualText (remove me) >> BDC\n" +
            "q BT /F1 1 Tf 3 Tr 50 700 Td ( ) Tj ET Q\n" +
            "BT /F2 20 Tf 1 0 0 rg 50 700 Td (remove me) Tj ET\n" +
            "EMC\n" + followingContent);
        var document = PdfDocument.Load(source);
        PdfTextMatch match = Assert.Single(document.Text.Find(
            "remove me", new PdfTextSearchOptions { MatchCase = true }));

        Assert.Throws<NotSupportedException>(() => document.Text.Replace(
            new PdfPageRegion(1, match.X, match.Y, match.Width, match.Height), "updated"));
    }
}
