using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    private const string ResidueHolderQuery = "TOPSECRET";

    [Fact]
    public void CompleteXrefProofAcceptsTablesLongerThanTheRecoveryCap() {
        const int objectCount = 110_000;
        var objects = new List<string>(objectCount) {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj"
        };
        for (int number = 4; number <= objectCount; number++) objects.Add(number + " 0 obj\nnull\nendobj");
        byte[] pdf = BuildPdf(objects, rootObjectNumber: 1);

        Assert.True(PdfSyntax.IsClassicRedactionOutput(pdf, PdfReadDocument.Open(pdf), default));
    }

    [Fact]
    [Trait("Category", "Performance")]
    public void ResidueSweepStaysLinearInStreamCount() {
        const int streamCount = 20_000;
        var objects = new List<string>(streamCount + 5) {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj"
        };
        byte[] data = Encoding.ASCII.GetBytes("plain note");
        for (int number = 4; number < 4 + streamCount; number++) objects.Add(BuildStreamObject(number, data));
        PdfDocument source = PdfDocument.Load(BuildPdf(objects, rootObjectNumber: 1));

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral(ResidueHolderQuery));
        stopwatch.Stop();

        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(10), "Residue sweep took " + stopwatch.Elapsed + ".");
    }

    [Theory]
    [InlineData("begincmap note TOPSECRET here", false)]
    [InlineData("begincmap 1 beginbfchar <41> <0041> endbfchar TOPSECRET endcmap", true)]
    public void WordBegincmapDoesNotHideStreamFromResidueSweep(string stream, bool referencedAsToUnicode) {
        string font = referencedAsToUnicode ? "/ToUnicode 6 0 R" : string.Empty;
        PdfDocument source = PdfDocument.Load(BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /PrivateData 6 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica " + font + " >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Public page) Tj ET")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(stream))
        }, rootObjectNumber: 1));

        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral(ResidueHolderQuery));

        Assert.Contains(plan.Findings, static finding => finding.Code == "redaction-search-unplanned-residue");
    }

    [Fact]
    public void FormStreamOverDecodedLimitBlocksSearchRedaction() {
        byte[] form = Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 700 Td (" + new string('x', 4096) + ") Tj ET");
        byte[] input = BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Fm0 6 0 R >> >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (TOPSECRET) Tj ET")),
            BuildStreamObject(6, form, "/Type /XObject /Subtype /Form /BBox [0 0 612 792]")
        }, rootObjectNumber: 1);
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxDecodedStreamBytes = 1024 } };
        PdfDocument source = PdfDocument.Load(input, options);
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral(ResidueHolderQuery));

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => source.Redactions.ApplyWithEvidence(plan));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
    }

    [Theory]
    [InlineData("% TOPSECRET\n", "")]
    [InlineData("", "TOPSECRET\n")]
    public void NoMatchSearchFindsSearchedTextOutsideParsedObjects(string betweenObjects, string afterEndOfFile) {
        byte[] input = BuildNoMatchPdf(betweenObjects, afterEndOfFile);
        PdfDocument source = PdfDocument.Load(input);

        PdfRedactionPlan held = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral(ResidueHolderQuery));
        PdfRedactionPlan absent = source.Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("absent"));
        PdfRedactionPlan numbers = source.Redactions.Search(new PdfRedactionSearchOptions().AddRegex(@"\d{3}"));

        Assert.Contains(held.Findings, static finding => finding.Code == "redaction-search-unplanned-residue");
        Assert.Throws<InvalidOperationException>(() => source.Redactions.ApplyWithEvidence(held));
        Assert.Equal(input, source.Redactions.ApplyWithEvidence(absent).Pdf);
        Assert.Equal(input, source.Redactions.ApplyWithEvidence(numbers).Pdf);
    }

    [Theory]
    [InlineData("/Span<</ActualText(TOPSECRET)>>BDC BT /F1 12 Tf 72 700 Td (TOPSECRET) Tj ET EMC")]
    [InlineData("/Figure<</Alt(TOPSECRET)>>BDC BT /F1 12 Tf 72 700 Td (visible) Tj ET EMC")]
    [InlineData("/Figure<</Alt(TOPSECRET)>>BDC 70 695 50 20 re f EMC")]
    [InlineData("/Figure<</Alt(TOPSECRET)>>BDC q 100 0 0 20 70 695 cm /Im0 Do Q EMC")]
    [InlineData("/Figure<</Alt(TOPSECRET)>>BDC /Fm0 Do EMC")]
    public void ManualAreaRemovalClearsMarkedContentOpenedOutsideTheTextObject(string markedContent) {
        PdfDocument source = PdfDocument.Load(BuildResidueHolderPdf(markedContent));
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, 250, 30) });

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

        string output = Encoding.GetEncoding("ISO-8859-1").GetString(result.Pdf);
        Assert.DoesNotContain(ResidueHolderQuery, output, StringComparison.Ordinal);
        Assert.Contains("/Span BMC", output, StringComparison.Ordinal);
        Assert.True(result.Evidence.IsVerified);
    }

    [Fact]
    public void VerificationRejectsAlternateTextLeftAroundRemovedContent() {
        const string markedContent = "/Span<</ActualText(TOPSECRET)>>BDC BT /F1 12 Tf 72 700 Td (TOPSECRET) Tj ET EMC";
        PdfDocument source = PdfDocument.Load(BuildResidueHolderPdf(markedContent));
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, 250, 30) });
        byte[] leftBehind = BuildResidueHolderPdf("/Span<</ActualText(TOPSECRET)>>BDC EMC");

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(leftBehind, plan, new PdfRedactionVerificationOptions());

        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactedContentOutsidePageContent" && issue.Marker == "AlternateText");
    }

    private static byte[] BuildNoMatchPdf(string betweenObjects, string afterEndOfFile) {
        byte[] pdf = BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n" + betweenObjects.TrimEnd(),
            BuildStreamObject(5, Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Public page) Tj ET"))
        }, rootObjectNumber: 1);
        return pdf.Concat(Encoding.ASCII.GetBytes(afterEndOfFile)).ToArray();
    }

    private static byte[] BuildResidueHolderPdf(string content) => BuildPdf(new[] {
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
        "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
        "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> /XObject << /Im0 6 0 R /Fm0 7 0 R >> >> >>\nendobj",
        "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
        BuildStreamObject(5, Encoding.ASCII.GetBytes(content + "\nBT /F1 12 Tf 72 500 Td (other text) Tj ET")),
        BuildStreamObject(6, new byte[] { 0x20, 0x40 }, "/Type /XObject /Subtype /Image /Width 2 /Height 1 /BitsPerComponent 8 /ColorSpace /DeviceGray"),
        BuildStreamObject(7, Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 700 Td (inside) Tj ET"), "/Type /XObject /Subtype /Form /BBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >>")
    }, rootObjectNumber: 1);
}
