using System;
using System.Text;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNavigationEncodingTests {
    [Fact]
    public void TextStringsPreserveCurrencyPunctuationAndNonbreakingSpace() {
        const string text = "€ “quoted” —\u00a0";
        string expected = "<FEFF" + BitConverter.ToString(Encoding.BigEndianUnicode.GetBytes(text)).Replace("-", "") + ">";
        Assert.Equal(expected, PdfSyntaxEscaper.TextString(text));
        var builder = new StringBuilder();
        PdfSyntaxEscaper.AppendTextStringCancellable(builder, text, CancellationToken.None);
        Assert.Equal(expected, builder.ToString());
        Assert.Equal(text, PdfTextString.Decode(PdfTextString.Encode(text)));
        Assert.Contains("/Contents " + expected, PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(0, 0, 10, 10, "https://example.com", text));
        Assert.Equal("•€—", PdfTextString.Decode(new byte[] { 0x80, 0xa0, 0x84 }));
    }

    [Theory]
    [InlineData("https://example.com/ż?q=é", "https://example.com/%C5%BC?q=%C3%A9")]
    [InlineData("https://żółć.pl/€", "https://xn--kda4b0koi.pl/%E2%82%AC")]
    [InlineData("assets/é.pdf#page", "assets/%C3%A9.pdf#page")]
    [InlineData("mailto:test@example.com?subject=é", "mailto:test@example.com?subject=%C3%A9")]
    public void UriActionsUseEscapedUriBytes(string uri, string escaped) {
        string expected = "(" + escaped + ")";
        Assert.Equal(expected, PdfSyntaxEscaper.UriString(uri));
        Assert.Contains("/URI " + expected, PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(0, 0, 10, 10, uri));
        Assert.Contains("/Base " + expected, PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, catalogUriBase: uri));
    }
}
