using System.Text;
using System.Net;
using System.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfReadbackCancellationOwnerTests {
    [Fact]
    public void FreeTextStyleAndAppearanceKeepValuesThroughCancellableScans() {
        using var cancellation = new CancellationTokenSource();
        string style = new string('x', 100_000) + ";font-size:12px;color:rgb(255,0,0);text-align:right";
        PdfFreeTextDefaultStyle parsed = PdfFreeTextStyleParser.ParseDefaultStyle(style, cancellation.Token);
        Assert.Equal(12D, parsed.FontSize);
        Assert.Equal(PdfAlign.Right, parsed.TextAlign);
        Assert.Equal(new PdfColor(1D, 0D, 0D), parsed.TextColor);

        string appearance = new string('x', 100_000) + " /F1 13 Tf 1 0 0 rg";
        Assert.True(PdfDefaultAppearanceParser.TryReadFontSize(appearance, out double fontSize, cancellation.Token));
        Assert.Equal(13D, fontSize);
        Assert.True(PdfDefaultAppearanceParser.TryReadTextColor(appearance, out PdfColor color, cancellation.Token));
        Assert.Equal(new PdfColor(1D, 0D, 0D), color);
    }

    [Fact]
    public void ReadbackUriAndOptionalContentChecksStayBounded() {
        using var cancellation = new CancellationTokenSource();
        Assert.True(Guard.IsUriAction("https://example.test/path", cancellation.Token));
        Assert.False(Guard.IsUriAction(new string('a', 66_000), cancellation.Token));
        Assert.True(Guard.IsNullOrWhiteSpaceCancellable(new string(' ', 100_000), cancellation.Token));
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => Guard.IsUriAction("https://example.test", cancellation.Token));
        Assert.ThrowsAny<OperationCanceledException>(() => Guard.IsNullOrWhiteSpaceCancellable(" ", cancellation.Token));
    }

    [Theory]
    [InlineData("A &amp; B &#x1F600;", "A & B \uD83D\uDE00")]
    [InlineData("&unknown; &copy; &amp", "&unknown; \u00A9 &amp")]
    [InlineData("&abc&copy;", "&abc\u00A9")]
    public void FreeTextPlainTextPreservesEntityDecoding(string source, string expected) {
        Assert.Equal(expected, PdfFreeTextStyleParser.ExtractPlainText(source));
        Assert.Equal(WebUtility.HtmlDecode(source), expected);
    }

    [Fact]
    public void FreeTextPlainTextPreservesLongMalformedEntityAndChecksCancellation() {
        string source = "&" + new string('a', 100_000) + "&amp;";
        Assert.Equal(WebUtility.HtmlDecode(source), PdfFreeTextStyleParser.ExtractPlainText(source));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfFreeTextStyleParser.ExtractPlainText(source, cancellation.Token));
    }

    [Fact]
    public void FreeTextPlainTextPreservesLongNumericEntityForms() {
        string[] sources = {
            "&#" + new string('0', 80) + "65;",
            "&#x" + new string('0', 80) + "41;",
            "&#" + new string(' ', 80) + "+65;",
            "&#65" + new string(' ', 80) + ";",
            "&#65" + new string('\0', 80) + ";",
            "&#" + new string('9', 80) + ";"
        };
        foreach (string source in sources) {
            Assert.Equal(WebUtility.HtmlDecode(source), PdfFreeTextStyleParser.ExtractPlainText(source));
        }
    }

    [Theory]
    [InlineData("\u03BB", "ASCII")]
    [InlineData("\u00E9", "\u03BB")]
    [InlineData("A", "\uD83D\uDE00")]
    public void LegacyPasswordPrefixPreservesFirst32EncodedBytes(string prefixCharacter, string suffix) {
        string password = string.Concat(Enumerable.Repeat(prefixCharacter, 40)) + suffix;
        byte[] expected = PdfWinAnsiEncoding.CanEncode(password, out _)
            ? PdfWinAnsiEncoding.Encode(password)
            : Encoding.UTF8.GetBytes(password);

        byte[] actual = PdfLegacyPasswordEncoding.EncodePrefix(password, CancellationToken.None);

        Assert.Equal(expected.Take(32), actual.Take(32));
        Assert.InRange(actual.Length, 1, 192);
    }

    [Fact]
    public void LegacyPasswordEncodingHonorsEncodingChoiceAfterConsumedPrefix() {
        string password = new string('\u00E9', 100_000) + "\u03BB";
        byte[] encoded = PdfLegacyPasswordEncoding.EncodePrefix(password, CancellationToken.None);

        Assert.Equal(Encoding.UTF8.GetBytes(new string('\u00E9', 16)), encoded.Take(32));
        using var source = new CancellationTokenSource();
        source.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfLegacyPasswordEncoding.EncodePrefix(password, source.Token));
    }

    [Fact]
    public void LegacySecurityWriterAndReaderAcceptLongPasswordsWithoutFullByteMaterialization() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Legacy password proof")).ToBytes();
        string userPassword = "\u03BB" + new string('x', 100_000);
        string ownerPassword = new string('\u00E9', 100_000) + "\u03BB";
        var encryption = new PdfStandardEncryptionOptions(userPassword) {
            OwnerPassword = ownerPassword,
            Algorithm = PdfStandardEncryptionAlgorithm.Aes128
        };

        byte[] encrypted = PdfSecurityEditor.Encrypt(source, encryption).Pdf;

        Assert.Equal(PdfPasswordAuthenticationRole.User,
            PdfInspector.Inspect(encrypted, new PdfLoadOptions { Password = userPassword }).Security.PasswordAuthenticationRole);
        Assert.Equal(PdfPasswordAuthenticationRole.Owner,
            PdfInspector.Inspect(encrypted, new PdfLoadOptions { Password = ownerPassword }).Security.PasswordAuthenticationRole);
    }

    [Fact]
    public void OrdinalReadbackComparisonKeepsOrderingAndHonorsCancellation() {
        string prefix = new string('x', 100_000);
        string left = prefix + "a";
        string right = prefix + "b";
        Assert.Equal(Math.Sign(StringComparer.Ordinal.Compare(left, right)),
            Math.Sign(PdfStringComparison.CompareOrdinal(left, right, CancellationToken.None)));
        using var source = new CancellationTokenSource();
        source.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfStringComparison.CompareOrdinal(left, right, source.Token));
    }

    [Fact]
    public void StreamDecoderResolvesIndirectFilterAndParameters() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfReference(1, 0);
        dictionary.Items["DecodeParms"] = new PdfReference(3, 0);
        var objects = new Dictionary<int, PdfIndirectObject> {
            [1] = new PdfIndirectObject(1, 0, new PdfReference(2, 0)),
            [2] = new PdfIndirectObject(2, 0, new PdfName("ASCIIHexDecode")),
            [3] = new PdfIndirectObject(3, 0, PdfNull.Instance)
        };

        Assert.Equal(new byte[] { 65, 66 }, StreamDecoder.DecodeRequired(dictionary, Encoding.ASCII.GetBytes("4142>"), objects));
        using var source = new CancellationTokenSource();
        source.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => StreamDecoder.DecodeRequired(
            dictionary, Encoding.ASCII.GetBytes("4142>"), objects, cancellationToken: source.Token));
    }

    [Fact]
    public void UnsupportedFilterReadbackResolvesIndirectArrayAndDeduplicatesLongNames() {
        string unsupportedName = new string('U', 100_000);
        var filters = new PdfArray();
        filters.Items.Add(new PdfName(unsupportedName));
        filters.Items.Add(new PdfName("FlateDecode"));
        filters.Items.Add(new PdfName(unsupportedName));
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfReference(1, 0);
        var objects = new Dictionary<int, PdfIndirectObject> {
            [1] = new PdfIndirectObject(1, 0, new PdfReference(2, 0)),
            [2] = new PdfIndirectObject(2, 0, filters)
        };

        Assert.Equal(new[] { unsupportedName },
            StreamDecoder.GetUnsupportedFiltersCancellable(dictionary, objects, CancellationToken.None));
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() =>
            StreamDecoder.GetUnsupportedFiltersCancellable(dictionary, objects, cancelled.Token));
    }
}
