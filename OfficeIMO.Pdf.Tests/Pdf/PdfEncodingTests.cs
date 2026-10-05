using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEncodingTests {
    [Theory]
    [InlineData("Earth’s € – ™", "4561727468907320A020852092")]
    [InlineData("˘ˇˆ˙˝˛˚˜", "18191A1B1C1D1E1F")]
    [InlineData("•†‡…—–ƒ⁄‹›−‰„“”‘’‚™ﬁﬂŁŒŠŸŽıłœšž", "808182838485868788898A8B8C8D8E8F909192939495969798999A9B9C9D9E")]
    public void PdfTextStringsUseDocumentEncodingRatherThanFontEncoding(string text, string hex) {
        Assert.Equal("<" + hex + ">", PdfSyntaxEscaper.TextString(text));
        var appended = new StringBuilder();
        PdfSyntaxEscaper.AppendTextStringCancellable(appended, text, CancellationToken.None);
        Assert.Equal("<" + hex + ">", appended.ToString());
        byte[] bytes = PdfTextString.DecodeHexBytes(hex);
        Assert.Equal(bytes, PdfTextString.Encode(text));
        Assert.Equal(text, PdfTextString.Decode(bytes));
        // Font glyph strings retain their separate Windows-1252 contract.
        Assert.Equal(new byte[] { 0x92, 0x80, 0x99 }, PdfWinAnsiEncoding.Encode("’€™"));
    }

    [Theory]
    [InlineData("\u00A0\u00AD", "FEFF00A000AD")]
    [InlineData("þÿ", "FEFF00FE00FF")]
    [InlineData("ÿþ", "FEFF00FF00FE")]
    [InlineData("ï»¿", "FEFF00EF00BB00BF")]
    public void PdfTextStringsUseUtf16ForUnrepresentedCharactersAndBomCollisions(string text, string hex) {
        Assert.Equal("<" + hex + ">", PdfSyntaxEscaper.TextString(text));
        var appended = new StringBuilder();
        PdfSyntaxEscaper.AppendTextStringCancellable(appended, text, CancellationToken.None);
        Assert.Equal("<" + hex + ">", appended.ToString());
        Assert.Equal(PdfTextString.DecodeHexBytes(hex), PdfTextString.Encode(text));
        Assert.Equal(text, PdfTextString.Decode(PdfTextString.Encode(text)));
    }

    [Fact]
    public void OutlineTitlesAndLinkDescriptionsAreDocumentTextRatherThanRawBytes() {
        string outline = PdfOutlineDictionaryBuilder.BuildOutlineItem("Earth’s €", 1, 0, 0, 0, 0, 0, 2, 100);
        Assert.Contains("/Title (Earth\u0090s \u00A0)", outline, StringComparison.Ordinal);
        string annotation = PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(
            10, 10, 20, 20, "https://example.org/", contents: "\u00A0\u00AD");
        Assert.Contains("/Contents <FEFF00A000AD>", annotation, StringComparison.Ordinal);
        Assert.Contains("/URI (https://example.org/)", annotation, StringComparison.Ordinal);
    }

    [Fact]
    public void Latin1DecodingPreservesEveryByteIncludingBinaryPdfContent() {
        byte[] bytes = Enumerable.Range(0, 256).Select(static value => (byte)value).ToArray();
        string expected = new string(bytes.Select(static value => (char)value).ToArray());

        Assert.Equal(expected, PdfEncoding.Latin1GetString(bytes));
        Assert.Equal(expected.Substring(31, 193), PdfEncoding.Latin1GetString(bytes, 31, 193));
        Assert.Equal(bytes, PdfEncoding.Latin1GetBytes(PdfEncoding.Latin1GetString(bytes)));
    }

    [Fact]
    public void CancellableStringMaterializationPreservesLargeLegacyTargetSlices() {
        string source = new string('a', 70_000) + "\u03A9" + new string('b', 70_000);
        char[] characters = source.ToCharArray();
        var builder = new StringBuilder("prefix").Append(source).Append("suffix");
        using var active = new CancellationTokenSource();
        CancellationToken token = active.Token;

        Assert.Equal(source, PdfEncoding.CharArrayToStringCancellable(characters, token));
        Assert.Equal(source, PdfEncoding.StringSliceCancellable("prefix" + source + "suffix", 6, source.Length, token));
        Assert.Equal(source, PdfEncoding.StringBuilderToStringCancellable(builder, 6, source.Length, token));

        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfEncoding.CharArrayToStringCancellable(characters, cancelled.Token));
    }
}
