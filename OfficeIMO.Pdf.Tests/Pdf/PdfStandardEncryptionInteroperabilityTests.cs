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

    // Acrobat X/XI pad /O and /U to 127 bytes; only the first 48 carry the hash and salts. Strict mode accepts them too,
    // because nothing is repaired: the defined bytes are used as stored.
    [Theory]
    [InlineData(PdfParsingMode.Lenient)]
    [InlineData(PdfParsingMode.Strict)]
    public void Revision6OwnerOnlyFilesWithPaddedPasswordEntriesOpenWithoutPasswordAndWithOwnerPassword(PdfParsingMode mode) {
        byte[] pdf = Fixture("r6-owner-only-padded-entries.pdf");

        foreach ((string? password, PdfPasswordAuthenticationRole role) in new[] {
            ((string?)null, PdfPasswordAuthenticationRole.User),
            ("owner", PdfPasswordAuthenticationRole.Owner)
        }) {
            var options = new PdfLoadOptions { Password = password, ParsingMode = mode };
            PdfDocumentInfo info = PdfInspector.Inspect(pdf, options);

            Assert.Equal("Padded entries R6", info.Metadata.Title);
            Assert.Equal(6, info.Security.EncryptionRevision);
            Assert.Equal(-1052, info.Security.EncryptionPermissions);
            Assert.Equal(role, info.Security.PasswordAuthenticationRole);
            Assert.Contains("Padded entries R6", Text(pdf, options), StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData("O")]
    [InlineData("U")]
    public void Revision6PasswordEntriesShorterThan48BytesAreRejected(string entry) {
        byte[] pdf = Fixture("r6-owner-only-padded-entries.pdf");
        System.Text.Encoding latin1 = System.Text.Encoding.GetEncoding(28591);
        string text = latin1.GetString(pdf);
        // Keep 47 bytes and blank the rest of the hex string so every object offset stays valid.
        text = System.Text.RegularExpressions.Regex.Replace(text, "/" + entry + @" <([0-9A-F]{254})>", match =>
            "/" + entry + " <" + match.Groups[1].Value.Substring(0, 94) + ">" + new string(' ', 254 - 94));
        pdf = latin1.GetBytes(text);

        PdfUnsupportedEncryptionException exception = Assert.Throws<PdfUnsupportedEncryptionException>(() =>
            PdfInspector.Inspect(pdf, new PdfLoadOptions { Password = "owner" }));
        Assert.Contains("invalid lengths", exception.Message, StringComparison.Ordinal);
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

    // An explicit /Identity filter and a named filter without /CFM (whose default is /None) both leave the stream as stored.
    [Theory]
    [InlineData("r6-identity-crypt-filter-stream.pdf", "Identity stream R6")]
    [InlineData("r6-crypt-filter-without-method.pdf", "Default method R6")]
    public void StreamsSelectingAnUnencryptedCryptFilterAreReadAsStored(string fixture, string text) {
        byte[] pdf = Fixture(fixture);
        var options = new PdfLoadOptions { Password = "open" };

        Assert.Contains(text, Text(pdf, options), StringComparison.Ordinal);
        Assert.Equal(text, PdfInspector.Inspect(pdf, options).Metadata.Title);
    }

    [Fact]
    public void DocumentsEncryptingOnlyEmbeddedFilesOpenWithoutPasswordAndKeepAttachmentsLocked() {
        byte[] pdf = Fixture("r6-embedded-files-only.pdf");

        Assert.Contains("Attachments only R6", Text(pdf, new PdfLoadOptions()), StringComparison.Ordinal);
        Assert.Throws<PdfPasswordRequiredException>(() => PdfDocument.Load(pdf).Attachments.Extract());
        var options = new PdfLoadOptions { Password = "open" };
        PdfExtractedAttachment attachment = Assert.Single(PdfDocument.Load(pdf, options).Attachments.Extract(options));
        Assert.Equal("Attachment secret", System.Text.Encoding.ASCII.GetString(attachment.Bytes));
    }

    [Theory]
    // /V 1 always uses 40 bits; /V 2 without /Length defaults to 40 bits, but some producers wrote 128-bit keys.
    [InlineData("r3-rc4-40-default-length.pdf", "RC4 40 R3")]
    [InlineData("r3-rc4-128-without-length.pdf", "RC4 128 R3")]
    public void Revision3WithoutLengthOpensWithTheDefaultOrAWrittenKeyLength(string fixture, string expected) {
        byte[] pdf = Fixture(fixture);

        Assert.Equal(expected, PdfInspector.Inspect(pdf, new PdfLoadOptions { Password = "open" }).Metadata.Title);
        // The fixtures restrict extraction for the user, so the text is read with the owner password.
        Assert.Contains(expected, Text(pdf, new PdfLoadOptions { Password = "owner" }), StringComparison.Ordinal);
    }

    [Fact]
    public void RewritesKeepEmbeddedFilesEncryptedWithTheirOwnCryptFilter() {
        var options = new PdfLoadOptions { Password = "open" };
        byte[] rewritten = PdfDocument.Load(Fixture("r6-embedded-files-only.pdf"), options).Pages.Extract("1").ToBytes();

        PdfExtractedAttachment attachment = Assert.Single(PdfDocument.Load(rewritten, options).Attachments.Extract(options));
        Assert.Equal("Attachment secret", System.Text.Encoding.ASCII.GetString(attachment.Bytes));
        // The rewrite keeps embedded-file-only encryption: pages read without the password, attachments do not.
        Assert.Contains("Attachments only R6", Text(rewritten, new PdfLoadOptions()), StringComparison.Ordinal);
        Assert.Throws<PdfPasswordRequiredException>(() => PdfDocument.Load(rewritten).Attachments.Extract());
    }

    [Fact]
    public void RewritesKeepALiteralFileIdentifierThatContainsABracket() {
        // The legacy file key depends on the first /ID string; a ']' inside the literal must not end the array.
        var options = new PdfLoadOptions { Password = "owner" };
        byte[] rewritten = PdfDocument.Load(Fixture("r3-literal-id-with-bracket.pdf"), options).Pages.Extract("1").ToBytes();

        Assert.Contains("Literal ID R3", Text(rewritten, options), StringComparison.Ordinal);
    }

    [Fact]
    public void IncrementalUpdatesKeepALiteralFileIdentifierWithHighBytes() {
        // The identifier holds ']' and the raw byte 0xE9; the appended trailer must repeat those exact bytes.
        var options = new PdfLoadOptions { Password = "owner" };
        PdfOperationResult<PdfDocument> rotation = PdfDocument.Load(Fixture("r3-literal-id-with-bracket.pdf"), options).Pages.RotateResult(90, "1");
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, rotation.MutationPlan!.ExecutionMode);
        byte[] rotated = rotation.RequireValue().ToBytes();

        Assert.Contains("Literal ID R3", Text(rotated, options), StringComparison.Ordinal);
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
