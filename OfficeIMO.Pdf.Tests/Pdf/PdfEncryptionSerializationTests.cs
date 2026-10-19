using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfEncryptionSerializationTests {
    private const string Content = "BT /F1 12 Tf 10 20 Td (Encrypted secret\nendobj\nendstream\nstream) Tj ET\n";

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void PublicEncryptionDoesNotTreatMetadataWordsInAStreamStringAsItsType(PdfStandardEncryptionAlgorithm algorithm) {
        const string description = "/Type /Metadata and /Length 999";
        byte[] source = CreatePdf("/Description (" + description + ") /Nested << /Type /Metadata >>", null);

        byte[] encrypted = PdfDocument.Load(source).Security.Encrypt(Options(algorithm)).Pdf;

        Assert.DoesNotContain("Encrypted secret", PdfEncoding.Latin1GetString(encrypted), StringComparison.Ordinal);
        Assert.Contains("Encrypted secret", PdfTextExtractor.ExtractAllText(encrypted, (PdfTextLayoutOptions?)null,
            new PdfLoadOptions { Password = "open" }), StringComparison.Ordinal);
        PdfStream stream = ReadContentStream(encrypted);
        Assert.Equal(description, stream.Dictionary.Get<PdfStringObj>("Description")!.Value);
        Assert.Equal(PdfEncoding.Latin1GetBytes(Content), stream.Data);
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void SerializedStringsAndStreamsCannotSupplyEncryptionBoundaries(PdfStandardEncryptionAlgorithm algorithm) {
        // Literal newlines are legal PDF syntax. The envelope is still produced by the real object serializer.
        const string description = "decoy\nstream\n/Length 999\nendstream\nendobj\n";
        string dictionary = "/Description (" + description + ") /Nested << /Length 1 >> " +
            "% an unmatched ( and fake /Length 2\n/Label <414243>";
        byte[] encrypted = CreatePdf(dictionary, Options(algorithm));

        PdfStream stream = ReadContentStream(encrypted);
        Assert.Equal(description, stream.Dictionary.Get<PdfStringObj>("Description")!.Value);
        Assert.Equal("ABC", stream.Dictionary.Get<PdfStringObj>("Label")!.Value);
        Assert.Equal(1, stream.Dictionary.Get<PdfDictionary>("Nested")!.Get<PdfNumber>("Length")!.Value);
        Assert.Equal(PdfEncoding.Latin1GetBytes(Content), stream.Data);
        Assert.DoesNotContain("Encrypted secret", PdfEncoding.Latin1GetString(encrypted), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void StreamLengthReplacementIgnoresLaterNestedAndStringLengths(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] encrypted = CreatePdf("/Description (/Length 999) /Nested << /Length 1 >>", Options(algorithm));

        PdfStream stream = ReadContentStream(encrypted);
        Assert.Equal("/Length 999", stream.Dictionary.Get<PdfStringObj>("Description")!.Value);
        Assert.Equal(1, stream.Dictionary.Get<PdfDictionary>("Nested")!.Get<PdfNumber>("Length")!.Value);
        Assert.Equal(PdfEncoding.Latin1GetBytes(Content), stream.Data);
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void ActualMetadataStreamsRemainExemptWhenRequested(PdfStandardEncryptionAlgorithm algorithm) {
        const string xml = "<metadata>visible metadata endobj endstream</metadata>";
        byte[] body = PdfObjectBytes.Concat(PdfEncoding.Latin1GetBytes("<< /Type /Metadata /Subtype /XML /Length " + xml.Length + " >>\nstream\n"),
            PdfEncoding.Latin1GetBytes(xml), PdfEncoding.Latin1GetBytes("\nendstream\n"));
        byte[] encrypted = PdfFileAssembler.Assemble(new[] {
            PdfObjectBytes.WrapIndirectObject(1, "<< /Type /Catalog /Pages 2 0 R /Metadata 4 0 R >>\n"),
            PdfObjectBytes.WrapIndirectObject(2, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>\n"),
            PdfObjectBytes.WrapIndirectObject(3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << >> >>\n"),
            PdfObjectBytes.WrapIndirectObject(4, body)
        }, 1, 0, encryption: Options(algorithm));

        Assert.Contains(xml, PdfEncoding.Latin1GetString(encrypted), StringComparison.Ordinal);
        var parsed = PdfSyntax.ParseObjects(encrypted, new PdfLoadOptions { Password = "open" });
        Assert.Equal(PdfEncoding.Latin1GetBytes(xml), Assert.IsType<PdfStream>(parsed.Map[4].Value).Data);
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void MismatchedStreamLengthFailsInsteadOfEncryptingGuessedBytes(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] stream = PdfObjectBytes.WrapIndirectObject(1, "<< /Length 1 >>\nstream\nsecret\nendstream\n");

        Assert.Throws<InvalidOperationException>(() => PdfStandardSecurityWriter.Encrypt(new[] { stream }, Options(algorithm)));
    }

    private static PdfStandardEncryptionOptions Options(PdfStandardEncryptionAlgorithm algorithm) => new("open") {
        OwnerPassword = "owner", Algorithm = algorithm, EncryptMetadata = false
    };

    private static PdfStream ReadContentStream(byte[] pdf) {
        var parsed = PdfSyntax.ParseObjects(pdf, new PdfLoadOptions { Password = "open" });
        return Assert.Single(parsed.Map.Values.Select(item => item.Value).OfType<PdfStream>());
    }

    private static byte[] CreatePdf(string extraDictionary, PdfStandardEncryptionOptions? encryption) {
        byte[] payload = PdfEncoding.Latin1GetBytes(Content);
        byte[] body = PdfObjectBytes.Concat(
            PdfEncoding.Latin1GetBytes("<< /Length " + payload.Length + " " + extraDictionary + " >>\nstream\n"),
            payload, PdfEncoding.Latin1GetBytes("\nendstream\n"));
        return PdfFileAssembler.Assemble(new[] {
            PdfObjectBytes.WrapIndirectObject(1, "<< /Type /Catalog /Pages 2 0 R >>\n"),
            PdfObjectBytes.WrapIndirectObject(2, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>\n"),
            PdfObjectBytes.WrapIndirectObject(3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\n"),
            PdfObjectBytes.WrapIndirectObject(4, body),
            PdfObjectBytes.WrapIndirectObject(5, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\n")
        }, 1, 0, encryption: encryption);
    }
}
