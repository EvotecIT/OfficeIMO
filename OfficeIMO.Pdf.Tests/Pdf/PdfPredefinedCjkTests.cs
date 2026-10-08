using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPredefinedCjkTests {
    [Fact]
    public void IndependentMismatchedEncodingAndCharacterCollectionRefusesSearch() {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures",
            "Interoperability", "PredefinedCjk", "mismatched-collection.pdf"));
        Assert.ThrowsAny<NotSupportedException>(() => PdfReadDocument.Open(source).ExtractText());
        PdfRedactionPlan plan = PdfDocument.Load(source).Redactions.Search(new PdfRedactionSearchOptions().AddLiteral("秘密"));
        Assert.False(plan.IsReviewable);
        Assert.Contains(plan.Findings, static finding => finding.Code == "RedactionSearchTextMappingUnsupported");
    }

    [Theory]
    [InlineData("japanese", "機密")]
    [InlineData("simplified-chinese", "秘密")]
    [InlineData("traditional-chinese", "秘密")]
    [InlineData("korean", "비밀")]
    public void IndependentPredefinedFontsDecodeMeasureAndRedactWithoutToUnicode(string name, string term) {
        string fixtures = Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "PredefinedCjk");
        byte[] original = File.ReadAllBytes(Path.Combine(fixtures, name + ".pdf"));
        using JsonDocument metadata = JsonDocument.Parse(File.ReadAllText(Path.Combine(fixtures, "source.json")));
        JsonElement expected = metadata.RootElement.GetProperty("cases").EnumerateArray().Single(
            item => item.GetProperty("file").GetString() == name + ".pdf");
        PdfReadDocument read = PdfReadDocument.Open(original);
        PdfTextSpan line = Assert.Single(read.Pages[0].GetTextSpans(), span => span.Text.Contains(term));
        Assert.Equal(expected.GetProperty("advancePoints").GetDouble(), line.Advance, 5);
        Assert.Contains(term, read.ExtractText());
        PdfDocument source = PdfDocument.Load(original);
        PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions {
            TextSelection = PdfRedactionTextSelection.MatchedGlyphs,
            ContentScope = PdfRedactionContentScope.TextOnly
        }.AddLiteral(term));
        Assert.True(plan.IsReviewable, string.Join("; ", plan.Findings.Select(static finding => finding.Message)));
        Assert.Single(plan.Areas);
        PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan).ThrowIfUnverified();
        string text = result.ToDocument().Read().Text;
        Assert.DoesNotContain(term, text);
        Assert.Contains("Before", text);
        Assert.Contains("after page", text);
        Assert.Contains("Public summary remains readable.", text);
        Assert.Equal(original, source.ToBytes());
    }

    [Theory]
    [InlineData("Other", "Japan1")]
    [InlineData("Adobe", "GB1")]
    public void IncompatibleCharacterCollectionsCannotSelectPredefinedMappings(string registry, string ordering) {
        Assert.Null(PdfPredefinedCMap.Find("UniJIS-UCS2-H", registry, ordering));
    }

    [Fact]
    public void PredefinedCodesEnforceDecodedOutputBudgetAndRejectSurrogates() {
        PdfPredefinedCMap mapping = PdfPredefinedCMap.Find("UniJIS-UCS2-H", "Adobe", "Japan1")!.Value;
        Assert.Throws<PdfReadLimitException>(() => mapping.TryDecode(Encoding.BigEndianUnicode.GetBytes("AB"), 1, out _));
        Assert.False(mapping.TryDecode(new byte[] { 0xD8, 0x00 }, 10, out _));
        Assert.False(mapping.TryDecode(new byte[] { 0x30 }, 10, out _));
    }

    [Fact]
    public void AdvertisedUnicodeOverridesPredefinedMapWithoutFallingBackForMissingCodes() {
        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes("1 beginbfchar <0041> <0058> endbfchar"), out ToUnicodeCMap? unicode));
        PdfFontResource font = new PdfFontResource("F1", "HeiseiMin-W3", "UniJIS-UCS2-H", true, unicode,
            fontSubtype: "Type0", predefinedCMap: PdfPredefinedCMap.Find("UniJIS-UCS2-H", "Adobe", "Japan1"));
        Func<byte[], int, string> decoder = ResourceResolver.CreateBudgetedDecoder(font);

        Assert.Equal("X", decoder(Encoding.BigEndianUnicode.GetBytes("A"), 10));
        Assert.ThrowsAny<NotSupportedException>(() => decoder(Encoding.BigEndianUnicode.GetBytes("AB"), 10));
    }

    [Theory]
    [InlineData("0009", "1 1 250", 4.5)]
    [InlineData("E000", "0 0 100", 1.8)]
    public void ExplicitUnicodeUsesPaintedCidWidthsIncludingNotdef(string code, string widths, double expectedAdvance) {
        byte[] source = BuildMappedTextSource("UniJIS-UCS2-H", "Japan1", code, widths, code + "> <0041");

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());

        Assert.Equal("A", span.Text);
        Assert.Equal(expectedAdvance, span.Advance, 5);
    }

    [Theory]
    [InlineData("0 65535 700", 12.6)]
    [InlineData("0 65535 700 4826 [900] 4825 4827 800 4826 [600]", 10.8)]
    [InlineData("4826 [900] 0 65535 700", 12.6)]
    [InlineData("0 65535 700 4825 4827 800", 14.4)]
    [InlineData("4825 4827 800 0 65535 700", 12.6)]
    [InlineData("4825 4827 800 4826 [600]", 10.8)]
    [InlineData("4826 [600] 4825 4827 800", 14.4)]
    public void CidWidthRangesRetainCompleteCoverageAndDeclarationPrecedence(string widths, double expectedAdvance) {
        byte[] source = BuildMappedTextSource("UniCNS-UCS2-H", "CNS1", "6A5F", widths);

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());

        Assert.Equal("機", span.Text);
        Assert.Equal(expectedAdvance, span.Advance, 5);
    }

    [Fact]
    public void IdentityFontsAlsoUseCompleteCidWidthRanges() {
        byte[] source = BuildMappedTextSource("Identity-H", "Japan1", "12DA", "0 65535 700", "12DA> <0041");

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());

        Assert.Equal("A", span.Text);
        Assert.Equal(12.6D, span.Advance, 5);
    }

    private static byte[] BuildMappedTextSource(string encoding, string ordering, string code, string widths, string? unicodePair = null) {
        string content = "BT /F1 18 Tf 72 700 Td <" + code + "> Tj ET";
        string unicode = "1 beginbfchar <" + unicodePair + "> endbfchar";
        string[] objects = {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type0 /BaseFont /HeiseiMin-W3 /Encoding /" + encoding + " /DescendantFonts [5 0 R] " + (unicodePair == null ? "" : "/ToUnicode 7 0 R ") + ">>",
            "<< /Type /Font /Subtype /CIDFontType0 /BaseFont /HeiseiMin-W3 /CIDSystemInfo << /Registry (Adobe) /Ordering (" + ordering + ") /Supplement 0 >> /DW 1000 /W [" + widths + "] >>",
            "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream",
            "<< /Length " + unicode.Length + " >>\nstream\n" + unicode + "\nendstream"
        };
        StringBuilder pdf = new StringBuilder("%PDF-1.7\n");
        List<int> offsets = new List<int> { 0 };
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(pdf.Length);
            pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = pdf.Length;
        pdf.Append("xref\n0 8\n0000000000 65535 f \n");
        foreach (int offset in offsets.Skip(1)) pdf.Append(offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        pdf.Append("trailer\n<< /Size 8 /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
