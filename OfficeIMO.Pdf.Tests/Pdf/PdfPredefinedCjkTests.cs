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
}
