using System.Text;
using System.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfReadbackCancellationOwnerTests {
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
}
