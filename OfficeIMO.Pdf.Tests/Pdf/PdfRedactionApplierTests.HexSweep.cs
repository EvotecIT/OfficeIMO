using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    [Theory]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs, false, false)]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs, true, false)]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks, false, false)]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks, true, false)]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs, false, true)]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs, true, true)]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks, false, true)]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks, true, true)]
    public void SearchSweepContinuesAfterMalformedHexAndPdfWhitespace(PdfRedactionTextSelection selection, bool retainedHexMatch, bool cmapHint) {
        byte[] stream = Encoding.ASCII.GetBytes((cmapHint ? "begincmap\n" : "") + "<not-hex> <4\0 1> <bad-token <736563726574>> <7 3 6 5 6 3 7 2 6 5 7 4>");
        byte[] input = BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /PrivateData 6 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Public page) Tj ET")),
            BuildStreamObject(6, stream)
        }, rootObjectNumber: 1);
        PdfDocument source = PdfDocument.Load(input);
        string query = retainedHexMatch ? "secret" : "absent";
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions { TextSelection = selection }.AddLiteral(query));
        Assert.Empty(plan.Areas);
        Assert.Equal(!retainedHexMatch, plan.IsReviewable);
        if (retainedHexMatch) {
            Assert.Contains(plan.Findings, static finding => finding.Code == "redaction-search-unplanned-residue");
            Assert.Throws<InvalidOperationException>(() => source.Redactions.ApplyWithEvidence(plan));
            Assert.False(source.Redactions.VerifyAppliedPlan(plan, new PdfRedactionVerificationOptions()).IsVerified);
        } else {
            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);
            Assert.True(result.Evidence.IsVerified);
            Assert.Equal(input, result.Pdf);
            using var reader = UglyToad.PdfPig.PdfDocument.Open(result.Pdf);
            Assert.Equal("Public page", reader.GetPage(1).Text);
        }
    }
}
