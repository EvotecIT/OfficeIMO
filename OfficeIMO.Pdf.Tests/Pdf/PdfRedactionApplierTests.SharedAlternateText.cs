using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    private const string TwinSections = "BT /F1 12 Tf /Span /P1 BDC 72 700 Td (visible) Tj EMC /Span /P1 BDC 0 -300 Td (kept) Tj EMC ET";
    private const string TwinFormSections = "/Span /P1 BDC BT /F1 12 Tf 72 700 Td (visible) Tj ET EMC /Span /P1 BDC BT /F1 12 Tf 72 300 Td (kept) Tj ET EMC";
    private const string TaggedSections = "/Figure <</MCID 0>> BDC BT /F1 12 Tf 72 700 Td (visible) Tj ET EMC /P <</MCID 1>> BDC BT /F1 12 Tf 72 300 Td (kept) Tj ET EMC";

    [Theory]
    [InlineData(TwinSections, "/Properties << /P1 6 0 R >>", "", "", "", "kept")]
    [InlineData("/Figure<</Alt(TOPSECRET)>>BDC BT /F1 12 Tf 72 700 Td (visible) Tj ET EMC /Figure<</Alt(TOPSECRET)>>BDC BT /F1 12 Tf 72 300 Td (kept) Tj ET EMC", "", "", "", "", "kept")]
    [InlineData("/Fm0 Do", "/XObject << /Fm0 7 0 R >>", TwinFormSections, "", "", "kept")]
    [InlineData("q 1 0 0 1 72 700 cm /Fm0 Do Q q 1 0 0 1 72 300 cm /Fm0 Do Q", "/XObject << /Fm0 7 0 R >>", "/Span /P1 BDC BT /F1 12 Tf 0 0 Td (visible) Tj ET EMC", "", "", "visible")]
    [InlineData(TaggedSections, "", "", "/K 0", "8 0 obj\n<< /Type /StructElem /S /P /P 9 0 R /Pg 3 0 R /K 1 >>\nendobj", "kept")]
    [InlineData(TaggedSections, "", "", "/K [0 1]", "8 0 obj\nnull\nendobj", "kept")]
    public void RemovedSectionTakesAlternateTextSharedWithRetainedContent(string content, string resources, string form, string structureKids, string secondElement, string retained) {
        byte[] input = BuildSharedAlternateTextPdf(content, resources, form, structureKids, secondElement);
        PdfDocument source = PdfDocument.Load(input);
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, 250, 30) });

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

        Assert.Contains(ResidueHolderQuery, Encoding.GetEncoding("ISO-8859-1").GetString(input), StringComparison.Ordinal);
        Assert.DoesNotContain(ResidueHolderQuery, Encoding.GetEncoding("ISO-8859-1").GetString(result.Pdf), StringComparison.Ordinal);
        Assert.Contains(retained, PdfDocument.Load(result.Pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(result.Evidence.IsVerified);
    }

    [Fact]
    public void VerificationRejectsRetainedSectionHoldingAlternateTextOfRemovedTwin() {
        PdfDocument source = PdfDocument.Load(BuildSharedAlternateTextPdf(TwinSections, "/Properties << /P1 6 0 R >>"));
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, 250, 30) });
        byte[] leftBehind = BuildSharedAlternateTextPdf("BT /F1 12 Tf /Span BMC 72 700 Td [-2778] TJ EMC /Span /P1 BDC 0 -300 Td (kept) Tj EMC ET", "/Properties << /P1 6 0 R >>");

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(leftBehind, plan, new PdfRedactionVerificationOptions());

        Assert.Contains(report.Issues, static issue => issue.Feature == "RedactedContentOutsidePageContent" && issue.Marker == "AlternateText");
    }

    [Fact]
    public void ManualAreaTakesSharedAlternateTextFromOtherPagesToo() {
        PdfDocument source = PdfDocument.Load(BuildTwoPageSharedListPdf("/Alt (TOPSECRET)"));
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { new PdfRedactionArea(1, 60, 690, 250, 30) });

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

        Assert.DoesNotContain(ResidueHolderQuery, Encoding.GetEncoding("ISO-8859-1").GetString(result.Pdf), StringComparison.Ordinal);
        Assert.True(result.Evidence.IsVerified);
    }

    [Theory]
    [InlineData("MATCHME note")]
    [InlineData("TOPSECRET")]
    public void PageScopedSearchEditsASharedListForItsOwnPageOnly(string description) {
        PdfDocument source = PdfDocument.Load(BuildTwoPageSharedListPdf("/Alt (" + description + ")"));
        var search = new PdfRedactionSearchOptions();
        search.AddLiteral("MATCHME");
        search.PageNumbers.Add(1);

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(source.Redactions.Search(search));

        PdfReadDocument output = PdfReadDocument.Open(result.Pdf);
        Assert.Null(ReadListAlternateText(output, 1));
        Assert.Equal(description, ReadListAlternateText(output, 2));
        Assert.True(result.Evidence.IsVerified);
    }

    [Fact]
    public void PageScopedSearchLeavesStructureElementOfThePageItDidNotCover() {
        byte[] input = BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 7 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 9 0 R >> >> >>\nendobj",
            "4 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 6 0 R /Resources << /Font << /F1 9 0 R >> >> >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 700 Td (MATCHME) Tj ET")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes("/P <</MCID 0>> BDC BT /F1 12 Tf 72 700 Td (other) Tj ET EMC")),
            "7 0 obj\n<< /Type /StructTreeRoot /K 8 0 R >>\nendobj",
            "8 0 obj\n<< /Type /StructElem /S /P /P 7 0 R /Pg 4 0 R /K 0 /Alt (MATCHME note) >>\nendobj",
            "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj"
        }, rootObjectNumber: 1);
        PdfDocument source = PdfDocument.Load(input);
        var search = new PdfRedactionSearchOptions();
        search.AddLiteral("MATCHME");
        search.PageNumbers.Add(1);

        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(source.Redactions.Search(search));

        Assert.Contains("/Alt (MATCHME note)", Encoding.GetEncoding("ISO-8859-1").GetString(result.Pdf), StringComparison.Ordinal);
        Assert.True(result.Evidence.IsVerified);
    }

    private static string? ReadListAlternateText(PdfReadDocument document, int pageNumber) {
        var page = (PdfDictionary)document.Objects[document.Pages[pageNumber - 1].ObjectNumber].Value;
        var resources = (PdfDictionary)PdfObjectLookup.Resolve(document.Objects, page.Items["Resources"])!;
        var properties = (PdfDictionary)PdfObjectLookup.Resolve(document.Objects, resources.Items["Properties"])!;
        var list = (PdfDictionary)PdfObjectLookup.Resolve(document.Objects, properties.Items["P1"])!;
        return list.Items.TryGetValue("Alt", out PdfObject? value) ? (PdfObjectLookup.Resolve(document.Objects, value) as PdfStringObj)?.Value : null;
    }

    private static byte[] BuildSharedAlternateTextPdf(string content, string resources, string form = "", string structureKids = "", string secondElement = "") {
        bool tagged = structureKids.Length > 0;
        var objects = new List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R " + (tagged ? "/StructTreeRoot 9 0 R" : string.Empty) + " >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> " + resources + " >> >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(content)),
            "6 0 obj\n<< /Alt (TOPSECRET) >>\nendobj"
        };
        if (form.Length > 0) objects.Add(BuildStreamObject(7, Encoding.ASCII.GetBytes(form), "/Type /XObject /Subtype /Form /BBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> /Properties << /P1 6 0 R >> >>"));
        if (tagged) {
            objects.Add(secondElement);
            objects.Add("9 0 obj\n<< /Type /StructTreeRoot /K [10 0 R " + (secondElement.Contains("StructElem") ? "8 0 R" : string.Empty) + "] >>\nendobj");
            objects.Add("10 0 obj\n<< /Type /StructElem /S /Figure /P 9 0 R /Pg 3 0 R " + structureKids + " /Alt (TOPSECRET) >>\nendobj");
        }
        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildTwoPageSharedListPdf(string list) => BuildPdf(new[] {
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
        "2 0 obj\n<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 612 792] >>\nendobj",
        "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources 7 0 R >>\nendobj",
        "4 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 6 0 R /Resources 7 0 R >>\nendobj",
        BuildStreamObject(5, Encoding.ASCII.GetBytes("/Span /P1 BDC BT /F1 12 Tf 72 700 Td (MATCHME) Tj ET EMC /Span /P1 BDC BT /F1 12 Tf 72 300 Td (kept) Tj ET EMC")),
        BuildStreamObject(6, Encoding.ASCII.GetBytes("/Span /P1 BDC BT /F1 12 Tf 72 700 Td (page two) Tj ET EMC")),
        "7 0 obj\n<< /Font << /F1 8 0 R >> /Properties << /P1 9 0 R >> >>\nendobj",
        "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
        "9 0 obj\n<< " + list + " >>\nendobj"
    }, rootObjectNumber: 1);
}
