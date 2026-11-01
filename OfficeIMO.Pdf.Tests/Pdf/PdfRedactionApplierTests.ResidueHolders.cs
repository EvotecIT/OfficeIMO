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

    [Fact]
    public void VerificationRejectsAlternateTextSectionThatLostPartOfItsContent() {
        const string survivor = "BT /F1 12 Tf 72 300 Td (kept) Tj ET";
        PdfDocument source = PdfDocument.Load(BuildResidueHolderPdf("/Figure<</Alt(TOPSECRET)>>BDC BT /F1 12 Tf 72 700 Td (visible) Tj ET " + survivor + " EMC"));
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, 250, 30) });
        byte[] leftBehind = BuildResidueHolderPdf("/Figure<</Alt(TOPSECRET)>>BDC " + survivor + " EMC");

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(leftBehind, plan, new PdfRedactionVerificationOptions());

        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactedContentOutsidePageContent" && issue.Marker == "AlternateText");
    }

    [Fact]
    public void VerificationAcceptsUnrelatedAlternateTextSectionThatWasAlreadyEmpty() {
        PdfDocument source = PdfDocument.Load(BuildResidueHolderPdf("BT /F1 12 Tf 72 700 Td (visible) Tj ET /Figure<</Alt(described)>>BDC EMC"));
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, 250, 30) });

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

        Assert.Contains("/Alt(described)", Encoding.GetEncoding("ISO-8859-1").GetString(result.Pdf), StringComparison.Ordinal);
        Assert.True(result.Evidence.IsVerified);
    }

    [Theory]
    [InlineData("70 695 50 20 re f", 250D)]
    [InlineData("q 100 0 0 20 70 695 cm /Im0 Do Q", 250D)]
    [InlineData("q 100 0 0 20 70 695 cm /Im0 Do Q", 30D)]
    [InlineData("/Fm1 Do", 250D)]
    public void RemovalClearsSectionsOpenedInAnotherContentStream(string paint, double areaWidth) {
        PdfDocument source = PdfDocument.Load(BuildResidueHolderPdf(new[] { "/Figure<</Alt(TOPSECRET)>>BDC", paint, "EMC" }));
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, areaWidth, 30) });

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

        Assert.DoesNotContain(ResidueHolderQuery, Encoding.GetEncoding("ISO-8859-1").GetString(result.Pdf), StringComparison.Ordinal);
        Assert.True(result.Evidence.IsVerified);
    }

    [Fact]
    public void ManualAreaRemovalClearsAlternateTextInNamedPropertyList() {
        byte[] input = BuildResidueHolderPdf("/Span /P1 BDC BT /F1 12 Tf 72 700 Td (visible) Tj ET EMC", "/Properties << /P1 8 0 R >>");
        PdfDocument source = PdfDocument.Load(input);
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, 250, 30) });

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

        Assert.DoesNotContain(ResidueHolderQuery, Encoding.GetEncoding("ISO-8859-1").GetString(result.Pdf), StringComparison.Ordinal);
        Assert.True(result.Evidence.IsVerified);
    }

    [Fact]
    public void ClearingOneSectionKeepsItsTwinInTheSameTextObject() {
        byte[] input = BuildResidueHolderPdf(
            "BT /F1 12 Tf /Span /P1 BDC 72 700 Td (visible) Tj EMC /Span /P1 BDC 0 -300 Td (kept) Tj EMC ET", "/Properties << /P1 9 0 R >>");
        PdfDocument source = PdfDocument.Load(input);
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, 250, 30) });

        string output = Encoding.GetEncoding("ISO-8859-1").GetString(source.Redactions.Apply(plan).ToBytes());

        Assert.True(output.IndexOf("/Span BMC", StringComparison.Ordinal) < output.IndexOf("/Span /P1 BDC", StringComparison.Ordinal));
        Assert.Contains("/Alt (described)", output, StringComparison.Ordinal);
    }

    [Fact]
    public void RedactionLeavesOptionalContentSectionsInPlace() {
        byte[] input = BuildResidueHolderPdf("/OC /P1 BDC BT /F1 12 Tf 72 700 Td (visible) Tj ET EMC", "/Properties << /P1 9 0 R >>");
        PdfDocument source = PdfDocument.Load(input);
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, 250, 30) });

        string output = Encoding.GetEncoding("ISO-8859-1").GetString(source.Redactions.Apply(plan).ToBytes());

        Assert.Contains("/OC /P1 BDC", output, StringComparison.Ordinal);
    }

    [Fact]
    public void PageScopedSearchLeavesSharedPropertyListWithThePageItDidNotCover() {
        const string section = "/Span /P1 BDC BT /F1 12 Tf 72 700 Td (TOPSECRET) Tj ET EMC";
        byte[] input = BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources 7 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 6 0 R /Resources 7 0 R >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(section)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(section)),
            "7 0 obj\n<< /Font << /F1 8 0 R >> /Properties << /P1 9 0 R >> >>\nendobj",
            "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "9 0 obj\n<< /ActualText (TOPSECRET) >>\nendobj"
        }, rootObjectNumber: 1);
        PdfDocument source = PdfDocument.Load(input);
        var search = new PdfRedactionSearchOptions();
        search.AddLiteral(ResidueHolderQuery);
        search.PageNumbers.Add(1);

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(source.Redactions.Search(search));

        string output = Encoding.GetEncoding("ISO-8859-1").GetString(result.Pdf);
        Assert.Contains("/Span BMC", output, StringComparison.Ordinal);
        Assert.Contains("/Span /P1 BDC", output, StringComparison.Ordinal);
        Assert.Contains("/ActualText (TOPSECRET)", output, StringComparison.Ordinal);
        Assert.True(result.Evidence.IsVerified);
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

    private static byte[] BuildResidueHolderPdf(string content, string resources = "") => BuildResidueHolderPdf(new[] { content }, resources);

    private static byte[] BuildResidueHolderPdf(string[] contents, string resources = "") {
        const int firstContent = 12;
        var objects = new List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents [" + string.Join(" ", contents.Select(static (_, index) => (firstContent + index) + " 0 R")) + " 5 0 R] " +
                "/Resources << /Font << /F1 4 0 R >> /XObject << /Im0 6 0 R /Fm0 7 0 R /Fm1 10 0 R >> " + resources + " >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 500 Td (other text) Tj ET")),
            BuildStreamObject(6, new byte[] { 0x20, 0x40 }, "/Type /XObject /Subtype /Image /Width 2 /Height 1 /BitsPerComponent 8 /ColorSpace /DeviceGray"),
            BuildStreamObject(7, Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 700 Td (inside) Tj ET"), "/Type /XObject /Subtype /Form /BBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >>"),
            "8 0 obj\n<< /ActualText (TOPSECRET) >>\nendobj",
            "9 0 obj\n<< /Type /OCG /Name (Layer) /Alt (described) >>\nendobj",
            BuildStreamObject(10, Encoding.ASCII.GetBytes("q 100 0 0 20 70 695 cm /Im0 Do Q"), "/Type /XObject /Subtype /Form /BBox [0 0 612 792] /Resources << /XObject << /Im0 11 0 R >> >>"),
            BuildStreamObject(11, new byte[] { 0x30, 0x50 }, "/Type /XObject /Subtype /Image /Width 2 /Height 1 /BitsPerComponent 8 /ColorSpace /DeviceGray")
        };
        objects.AddRange(contents.Select(static (content, index) => BuildStreamObject(firstContent + index, Encoding.ASCII.GetBytes(content))));
        return BuildPdf(objects.ToArray(), rootObjectNumber: 1);
    }
}
