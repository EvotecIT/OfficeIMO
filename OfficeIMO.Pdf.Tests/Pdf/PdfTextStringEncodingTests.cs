using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextStringEncodingTests {
    [Fact]
    public void ImportedPdfDocEncodingPunctuationUsesThePdfMapping() {
        Assert.Equal("—–•€Ł", PdfTextString.Decode([0x84, 0x85, 0x80, 0xA0, 0x95]));
    }

    [Theory]
    [InlineData("Review — approved")]
    [InlineData("café · Zażółć · 東京")]
    [InlineData("\u0018")]
    public void NonAsciiTextStringsUseExplicitUnicodeForIndependentReaders(string text) {
        byte[] encoded = PdfTextString.Encode(text);
        Assert.Equal(new byte[] { 0xFE, 0xFF }, encoded.Take(2));
        Assert.Equal(text, System.Text.Encoding.BigEndianUnicode.GetString(encoded, 2, encoded.Length - 2));
    }
}
