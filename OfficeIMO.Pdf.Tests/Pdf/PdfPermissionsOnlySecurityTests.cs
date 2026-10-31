using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPermissionsOnlySecurityTests {
    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void WriterOpensWithoutPasswordAndKeepsPermissionChangesOwnerAuthorized(PdfStandardEncryptionAlgorithm algorithm) {
        var encryption = PdfStandardEncryptionOptions.ForPermissions("owner-only", PdfStandardPermissions.Print);
        encryption.Algorithm = algorithm;
        var copy = encryption.Clone();
        Assert.Empty(copy.UserPassword);
        Assert.Equal("owner-only", copy.OwnerPassword);
        Assert.True(copy.HasDistinctOwnerPassword);
        Assert.Throws<ArgumentException>(() => copy.OwnerPassword = null);
        Assert.Throws<ArgumentException>(() => PdfStandardEncryptionOptions.ForPermissions("", PdfStandardPermissions.Print));

        using var document = PdfDocument.Create(new PdfCreateOptions { EmbedDefaultFonts = false }.SetEncryption(copy))
            .Paragraph(paragraph => paragraph.Text("Permissions-only proof"));
        byte[] bytes = document.ToBytes();
        var reader = PdfInspector.Inspect(bytes);
        Assert.True(reader.Security.HasEncryption);
        Assert.Equal(PdfPasswordAuthenticationRole.User, reader.Security.PasswordAuthenticationRole);
        Assert.Equal(PdfStandardPermissions.Print, reader.Security.AllowedStandardPermissions);
        Assert.Throws<PdfPermissionDeniedException>(() => PdfTextExtractor.ExtractAllText(bytes));

        var owner = new PdfLoadOptions { Password = "owner-only" };
        Assert.Equal(PdfPasswordAuthenticationRole.Owner, PdfInspector.Inspect(bytes, owner).Security.PasswordAuthenticationRole);
        Assert.Contains("Permissions-only proof", PdfTextExtractor.ExtractAllText(bytes, (PdfTextLayoutOptions?)null, owner), StringComparison.Ordinal);
        using var opened = PdfDocument.Load(bytes);
        Assert.False(opened.DecryptResult("").Succeeded);
        Assert.False(PdfSecurityEditor.Decrypt(bytes, "owner-only").IsEncrypted);
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256, "\u00e9", "e\u0301", false)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128, "12345678901234567890123456789012-open", "12345678901234567890123456789012-owner", false)]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4, "12345678901234567890123456789012-open", "12345678901234567890123456789012-owner", false)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256, "open-secret", "owner-secret", true)]
    public void DistinctOwnerAccessUsesTheSelectedAlgorithmsEffectiveCredentials(
        PdfStandardEncryptionAlgorithm algorithm, string user, string owner, bool distinct) {
        var options = new PdfStandardEncryptionOptions(user) { OwnerPassword = owner, Algorithm = algorithm };
        Assert.Equal(distinct, options.HasDistinctOwnerPassword);
        using var document = PdfDocument.Create(new PdfCreateOptions { EmbedDefaultFonts = false }.SetEncryption(options));
        document.Content.Paragraph(paragraph => paragraph.Text("Password equivalence proof"));
        var opened = PdfInspector.Inspect(document.ToBytes(), new PdfLoadOptions { Password = user });
        Assert.Equal(distinct ? PdfPasswordAuthenticationRole.User : PdfPasswordAuthenticationRole.Owner,
            opened.Security.PasswordAuthenticationRole);
    }

    [Fact]
    public void ExistingDocumentRewritePreservesContentWithAnEmptyOpenPassword() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        var encryption = PdfStandardEncryptionOptions.ForPermissions("rewrite-owner", PdfStandardPermissions.Print);
        var result = PdfSecurityEditor.Encrypt(source, encryption);
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
        var opened = PdfInspector.Inspect(result.Pdf);
        Assert.Equal(result.PreservationReport.Original.PageCount, opened.PageCount);
        Assert.Equal(PdfPasswordAuthenticationRole.User, opened.Security.PasswordAuthenticationRole);
        Assert.Equal(PdfStandardPermissions.Print, opened.Security.AllowedStandardPermissions);
        var restored = PdfSecurityEditor.Decrypt(result.Pdf, "rewrite-owner");
        Assert.True(restored.PreservationReport.IsPreserved, restored.PreservationReport.Summary);
        Assert.False(restored.IsEncrypted);
    }
}
