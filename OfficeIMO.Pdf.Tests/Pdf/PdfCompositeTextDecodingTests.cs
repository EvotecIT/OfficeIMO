using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCompositeTextDecodingTests {
    [Fact]
    public void IndependentCompositeFontWithoutUnicodeMappingRejectsTextExtraction() {
        byte[] source = LoadIndependentSource();

        NotSupportedException error = Assert.ThrowsAny<NotSupportedException>(
            () => PdfReadDocument.Open(source).ExtractText());

        Assert.Contains("UniJIS-UCS2-H", error.Message);
    }

    [Theory]
    [InlineData(PdfRedactionTextSelection.LogicalBlocks)]
    [InlineData(PdfRedactionTextSelection.MatchedGlyphs)]
    public void IndependentCompositeFontSearchReportsBlockedMappingInsteadOfNoMatches(PdfRedactionTextSelection selection) {
        PdfDocument document = PdfDocument.Load(LoadIndependentSource());

        PdfRedactionPlan plan = document.Redactions.Search(new PdfRedactionSearchOptions {
            TextSelection = selection
        }.AddLiteral("private account 123"));

        Assert.False(plan.IsReviewable);
        Assert.Empty(plan.Areas);
        PdfDiagnosticFinding finding = Assert.Single(plan.Findings,
            static finding => finding.Code == "RedactionSearchTextMappingUnsupported");
        Assert.Equal(PdfDiagnosticSeverity.Error, finding.Severity);
        Assert.Contains("UniJIS-UCS2-H", finding.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IndependentCompositeFontBlocksManualAndVerificationPlanning(bool verification) {
        byte[] source = LoadIndependentSource();
        PdfRedactionArea[] areas = new[] { new PdfRedactionArea(1, 50, 680, 200, 40) };

        PdfRedactionPlan plan = verification
            ? PdfRedactionPlanner.PlanForVerification(source, areas, options: null)
            : PdfDocument.Load(source).Redactions.Plan(areas);

        Assert.False(plan.IsReviewable);
        Assert.Contains(plan.Findings, static finding => finding.Severity == PdfDiagnosticSeverity.Error &&
            finding.Message.Contains("UniJIS-UCS2-H"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AdvertisedCompositeUnicodeMapMustContainMappings(bool emptyMap) {
        ToUnicodeCMap? map = null;
        if (emptyMap) {
            Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes("begincmap endcmap"), out map));
        }
        PdfFontResource font = new PdfFontResource("F1", "HeiseiMin-W3", "UniJIS-UCS2-H", true, map, fontSubtype: "Type0");
        Func<byte[], int, string> decode = ResourceResolver.CreateBudgetedDecoder(font);

        Assert.ThrowsAny<NotSupportedException>(() => decode(new byte[] { 0, 65 }, 10));
        Assert.Equal(string.Empty, decode(Array.Empty<byte>(), 10));
    }

    [Fact]
    public void NamedCompositeFontWithUnicodeMappingDecodesMappedText() {
        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes("1 beginbfchar <0041> <0041> endbfchar"), out ToUnicodeCMap? map));
        PdfFontResource font = new PdfFontResource("F1", "HeiseiMin-W3", "UniJIS-UCS2-H", true, map, fontSubtype: "Type0");

        Assert.Equal("A", ResourceResolver.CreateBudgetedDecoder(font)(new byte[] { 0, 65 }, 10));
    }

    [Fact]
    public void UnusedCompositeFontDoesNotBlockKnownText() {
        const string composite = "/Unused << /Type /Font /Subtype /Type0 /BaseFont /HeiseiMin-W3 " +
            "/Encoding /UniJIS-UCS2-H /DescendantFonts [<< /Type /Font /Subtype /CIDFontType0 " +
            "/BaseFont /HeiseiMin-W3 /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 5 >> >>] >>";
        byte[] original = PdfReaderAndFooterRegressionTests.BuildSingleStreamPdf("BT /F1 12 Tf 72 700 Td (Keep me public) Tj ET");
        byte[] source = Encoding.ASCII.GetBytes(Encoding.ASCII.GetString(original)
            .Replace("/Font << /F1 4 0 R >>", "/Font << /F1 4 0 R " + composite + " >>"));

        Assert.Contains("Keep me public", PdfReadDocument.Open(source).ExtractText());
        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Search(
            new PdfRedactionSearchOptions().AddLiteral("Keep"));

        Assert.True(plan.IsReviewable);
        Assert.Single(plan.Areas);
    }

    private static byte[] LoadIndependentSource() => File.ReadAllBytes(Path.Combine(
        AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "CompositeTextDecoding", "unijis-without-to-unicode.pdf"));
}
