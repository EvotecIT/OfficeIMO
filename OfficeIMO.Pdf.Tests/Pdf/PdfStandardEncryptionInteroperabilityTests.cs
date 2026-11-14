using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Standard security files written by New-StandardEncryptionFixtures.ps1 with .NET cryptography alone
/// (see Fixtures/Interoperability/StandardEncryption/SOURCE.md).
/// </summary>
public sealed class PdfStandardEncryptionInteroperabilityTests {
    private static byte[] Fixture(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "StandardEncryption", name));

    private static string Text(byte[] pdf, PdfLoadOptions options) => PdfDocument.Load(pdf, options).Reader.Text();

    [Theory]
    // RFC 4013 maps U+00AA to "a" and removes U+00AD; the file stores "SaSLprep".
    [InlineData("r6-saslprep.pdf", "S\u00AASL\u00ADprep", "SASLprep R6")]
    // The producer skipped SASLprep and stored the soft hyphen; the NFKC fallback still authenticates.
    [InlineData("r6-unprepared-password.pdf", "pass\u00ADword", "Unprepared R6")]
    public void Revision6PasswordsAreComparedAfterSaslPrep(string fixture, string password, string expected) {
        byte[] pdf = Fixture(fixture);

        PdfDocumentInfo info = PdfInspector.Inspect(pdf, new PdfLoadOptions { Password = password });

        Assert.Equal(expected, info.Metadata.Title);
        Assert.Equal(PdfPasswordAuthenticationRole.User, info.Security.PasswordAuthenticationRole);
        Assert.Throws<PdfInvalidPasswordException>(() => PdfInspector.Inspect(pdf, new PdfLoadOptions { Password = "SaSL-prep" }));
    }

    [Fact]
    public void SaslPrepPasswordsWrittenByOfficeIMOOpenWithEquivalentInput() {
        byte[] pdf = PdfDocument.Create(new PdfCreateOptions().SetEncryption(new PdfStandardEncryptionOptions("S\u00AASL\u00ADprep\u00A0two")))
            .Paragraph(paragraph => paragraph.AddText("Prepared"))
            .ToBytes();

        // SASLprep stores "SaSLprep two": the no-break space maps to SPACE and the soft hyphen to nothing.
        Assert.Contains("Prepared", Text(pdf, new PdfLoadOptions { Password = "SaSLprep two" }), StringComparison.Ordinal);
    }

    public static TheoryData<IOfficeAesCryptographyProvider?> AesProviders => new() {
        null,
        // The caller-supplied provider path computes AES-GCM with the managed construction on every target.
        new OfficeManagedAesCryptographyProvider()
    };

    [Theory]
    [MemberData(nameof(AesProviders))]
    public void Revision7AesGcmDocumentsOpenWithUserAndOwnerPasswords(IOfficeAesCryptographyProvider? provider) {
        byte[] pdf = Fixture("r7-aes-gcm.pdf");

        foreach ((string password, PdfPasswordAuthenticationRole role) in new[] {
            ("open", PdfPasswordAuthenticationRole.User),
            ("owner", PdfPasswordAuthenticationRole.Owner)
        }) {
            var options = new PdfLoadOptions { Password = password, AesCryptographyProvider = provider };
            PdfDocumentInfo info = PdfInspector.Inspect(pdf, options);

            Assert.Equal("AES-GCM R7", info.Metadata.Title);
            Assert.Equal(7, info.Security.EncryptionRevision);
            Assert.Equal(6, info.Security.EncryptionVersion);
            Assert.Equal(role, info.Security.PasswordAuthenticationRole);
            Assert.Contains("AES-GCM R7", Text(pdf, options), StringComparison.Ordinal);
        }

        Assert.Throws<PdfInvalidPasswordException>(() => PdfInspector.Inspect(pdf, new PdfLoadOptions { Password = "wrong" }));
        Assert.Throws<PdfPasswordRequiredException>(() => PdfInspector.Inspect(pdf));
    }

    [Theory]
    [MemberData(nameof(AesProviders))]
    public void Revision7AesGcmRejectsModifiedCiphertext(IOfficeAesCryptographyProvider? provider) {
        byte[] pdf = Fixture("r7-aes-gcm.pdf");
        // Flip one ciphertext byte of the page content stream (after its 12-byte nonce).
        int stream = IndexOf(pdf, "stream\n") + "stream\n".Length;
        pdf[stream + 12] ^= 0x01;

        Assert.Throws<PdfEncryptedContentIntegrityException>(() =>
            Text(pdf, new PdfLoadOptions { Password = "open", AesCryptographyProvider = provider }));
    }

    [Fact]
    public void StreamsSelectingTheIdentityCryptFilterAreReadAsStored() {
        byte[] pdf = Fixture("r6-identity-crypt-filter-stream.pdf");
        var options = new PdfLoadOptions { Password = "open" };

        Assert.Contains("Identity stream R6", Text(pdf, options), StringComparison.Ordinal);
        Assert.Equal("Identity stream R6", PdfInspector.Inspect(pdf, options).Metadata.Title);
    }

    [Fact]
    public void DocumentsEncryptingOnlyEmbeddedFilesOpenWithoutPasswordAndKeepAttachmentsLocked() {
        byte[] pdf = Fixture("r6-embedded-files-only.pdf");

        Assert.Contains("Attachments only R6", Text(pdf, new PdfLoadOptions()), StringComparison.Ordinal);
        var options = new PdfLoadOptions { Password = "open" };
        PdfExtractedAttachment attachment = Assert.Single(PdfDocument.Load(pdf, options).Attachments.Extract(options));
        Assert.Equal("Attachment secret", System.Text.Encoding.ASCII.GetString(attachment.Bytes));
    }

    [Fact]
    public void Revision3WithoutLengthUsesTheFortyBitDefault() {
        byte[] pdf = Fixture("r3-rc4-40-default-length.pdf");
        // The fixture restricts extraction for the user, so the text is read with the owner password.
        Assert.Equal("RC4 40 R3", PdfInspector.Inspect(pdf, new PdfLoadOptions { Password = "open" }).Metadata.Title);
        Assert.Contains("RC4 40 R3", Text(pdf, new PdfLoadOptions { Password = "owner" }), StringComparison.Ordinal);
    }

    private static int IndexOf(byte[] data, string marker) {
        byte[] pattern = System.Text.Encoding.ASCII.GetBytes(marker);
        for (int index = 0; index <= data.Length - pattern.Length; index++) {
            int match = 0;
            while (match < pattern.Length && data[index + match] == pattern[match]) match++;
            if (match == pattern.Length) return index;
        }

        throw new InvalidOperationException("Marker not found.");
    }
}
