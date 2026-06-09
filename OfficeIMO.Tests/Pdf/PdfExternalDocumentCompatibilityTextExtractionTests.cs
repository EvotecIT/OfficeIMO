using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {

    [Fact]
    public void ExtractAllText_PreservesSpacingFromExternalTjArrays() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n[(External) -360 (PDF) -360 (text)] TJ\nET\n");

        string allText = PdfTextExtractor.ExtractAllText(pdf);
        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("External PDF text", Normalize(allText), StringComparison.Ordinal);
        Assert.Contains("External PDF text", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesExternalToUnicodeCMap() {
        byte[] pdf = BuildExternalToUnicodePdf();

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("Zed", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesExternalToUnicodeBfRangeArrayCMap() {
        byte[] pdf = BuildExternalToUnicodeBfRangeArrayPdf();

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("Zfi!", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ToUnicodeCMap_TryEncodeTextGreedilyMatchesMultiScalarEntries() {
        const string cmap = "beginbfchar\n<01> <0066>\n<02> <00660069>\n<03> <0069>\nendbfchar";

        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(cmap), out ToUnicodeCMap? parsed));
        Assert.NotNull(parsed);

        Assert.True(parsed!.TryEncodeText("fi", out string ligatureHex));
        Assert.Equal("02", ligatureHex);
        Assert.True(parsed.TryEncodeText("fii", out string mixedHex));
        Assert.Equal("0203", mixedHex);
    }

    [Fact]
    public void ExtractTextByPage_UsesExternalThreeByteToUnicodeCMap() {
        byte[] pdf = BuildExternalThreeByteToUnicodePdf();

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("Z", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesSimpleFontMacRomanEncoding() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n<4361668E> Tj\nET\n",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /MacRomanEncoding >>");

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("Caf\u00E9", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesSimpleFontEncodingDifferencesWithMacRomanBase() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n<8E20C8> Tj\nET\n",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /BaseEncoding /MacRomanEncoding /Differences [ 200 /Euro ] >> >>");

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("\u00E9 \u20AC", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesSimpleFontStandardEncoding() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n<27206020A920AEAFE1F1> Tj\nET\n",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /StandardEncoding >>");

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));
        string normalized = Normalize(pageText);

        Assert.Contains("\u2019 \u2018 '", normalized, StringComparison.Ordinal);
        Assert.Contains("fifl\u00C6\u00E6", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesSimpleFontEncodingDifferencesWithStandardBase() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n<2720AE20AF> Tj\nET\n",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /BaseEncoding /StandardEncoding /Differences [ 174 /Euro ] >> >>");

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("\u2019 \u20AC fl", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesSimpleFontEncodingDifferencesWithLatinGlyphNames() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n<414243444546474849> Tj\nET\n",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /BaseEncoding /WinAnsiEncoding /Differences [ 65 /Agrave /eacute /ccedilla /ntilde /germandbls /Oslash /questiondown /Aacute.alt /ydieresis ] >> >>");

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("\u00C0\u00E9\u00E7\u00F1\u00DF\u00D8\u00BF\u00C1\u00FF", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesSimpleFontEncodingDifferencesWithCompositeGlyphNames() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n<414243> Tj\nET\n",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /BaseEncoding /WinAnsiEncoding /Differences [ 65 /f_f_i /A_uni0301 /f_f_l.alt ] >> >>");

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("ffiA\u0301ffl", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesSimpleFontEncodingDifferencesWithCommonWinAnsiGlyphNames() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n<4142434445464748494A4B4C4D4E4F> Tj\nET\n",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /BaseEncoding /WinAnsiEncoding /Differences [ 65 /OE /oe /Scaron /scaron /Zcaron /zcaron /Ydieresis /florin /perthousand /quotesinglbase /quotedblbase /guilsinglleft /guilsinglright /dagger /circumflex ] >> >>");

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("\u0152\u0153\u0160\u0161\u017D\u017E\u0178\u0192\u2030\u201A\u201E\u2039\u203A\u2020\u02C6", Normalize(pageText), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_UsesBuiltInStandardEncodingForStandardType1Font() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n<2720AEAFE1F1> Tj\nET\n",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));
        string normalized = Normalize(pageText);

        Assert.Contains("\u2019", normalized, StringComparison.Ordinal);
        Assert.Contains("fifl\u00C6\u00E6", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextByPage_KeepsWinAnsiFallbackForUnknownSimpleFontWithoutEncoding() {
        byte[] pdf = BuildExternalSinglePagePdf(
            "BT\n/F13 12 Tf\n72 720 Td\n<436166E9> Tj\nET\n",
            "<< /Type /Font /Subtype /Type1 /BaseFont /CustomSans >>");

        string pageText = Assert.Single(PdfTextExtractor.ExtractTextByPage(pdf));

        Assert.Contains("Caf\u00E9", Normalize(pageText), StringComparison.Ordinal);
    }

}
