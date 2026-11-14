using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Org.BouncyCastle.Cms;
using Xunit;
using Fixture = OfficeIMO.Tests.Pdf.PublicKeyEncryptedPdfFixture;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPublicKeyEncryptedReadTests {
    // RSA key generation dominates the run time, so the read-only recipients are shared by every test.
    private static readonly X509Certificate2 _recipient = Fixture.CreateRecipient("OfficeIMO PDF recipient");
    private static readonly X509Certificate2 _otherRecipient = Fixture.CreateRecipient("OfficeIMO other recipient");

    public static TheoryData<string, int, Fixture.Cipher, int, bool, string> Variants => new() {
        { "adbe.pkcs7.s3", 1, Fixture.Cipher.Rc4, 40, true, CmsEnvelopedGenerator.Aes256Cbc },
        { "adbe.pkcs7.s4", 2, Fixture.Cipher.Rc4, 128, true, CmsEnvelopedGenerator.Aes128Cbc },
        { "adbe.pkcs7.s4", 2, Fixture.Cipher.Rc4, 128, true, CmsEnvelopedGenerator.DesEde3Cbc },
        { "adbe.pkcs7.s5", 4, Fixture.Cipher.Rc4, 128, true, CmsEnvelopedGenerator.Aes256Cbc },
        { "adbe.pkcs7.s5", 4, Fixture.Cipher.AesV2, 128, true, CmsEnvelopedGenerator.Aes256Cbc },
        { "adbe.pkcs7.s5", 4, Fixture.Cipher.AesV2, 128, false, CmsEnvelopedGenerator.Aes256Cbc },
        { "adbe.pkcs7.s5", 5, Fixture.Cipher.AesV3, 256, true, CmsEnvelopedGenerator.Aes256Cbc },
        { "adbe.pkcs7.s5", 5, Fixture.Cipher.AesV3, 256, false, CmsEnvelopedGenerator.Aes256Cbc }
    };

    [Theory]
    [MemberData(nameof(Variants))]
    public void OpensEachSubFilterAndCipherWithTheRecipientCertificate(
        string subFilter, int version, Fixture.Cipher method, int keyLengthBits, bool encryptMetadata, string envelopeAlgorithm) {
        int permissions = Fixture.AllPermissions & ~(1 << 2);
        byte[] pdf = Fixture.Create(new Fixture.Settings {
            SubFilter = subFilter,
            Version = version,
            Method = method,
            KeyLengthBits = keyLengthBits,
            EncryptMetadata = encryptMetadata,
            EnvelopeAlgorithm = envelopeAlgorithm,
            Permissions = permissions
        }, _recipient, _otherRecipient);
        PdfLoadOptions options = OptionsFor(_recipient);

        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, options);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf, options);

        Assert.Contains(Fixture.Text, text, StringComparison.Ordinal);
        Assert.Equal(Fixture.Title, info.Metadata.Title);
        Assert.Equal("Adobe.PubSec", info.Security.EncryptionFilter);
        Assert.Equal(subFilter, info.Security.EncryptionSubFilter);
        Assert.Equal(PdfPasswordAuthenticationRole.PublicKeyRecipient, info.Security.PasswordAuthenticationRole);
        Assert.False(info.Security.HasOwnerAuthorization);
        Assert.Equal(permissions, info.Security.EncryptionPermissions);
        Assert.False(info.Security.AllowsPrinting);
        Assert.True(info.Security.AllowsCopying);
    }

    [Fact]
    public void SingleRecipientStringAndIdentityStringFilterAreAccepted() {
        byte[] pdf = Fixture.Create(new Fixture.Settings {
            Method = Fixture.Cipher.AesV2,
            Version = 4,
            KeyLengthBits = 128,
            SingleRecipientString = true,
            StringFilterOverride = "Identity"
        }, _recipient);

        PdfDocumentInfo info = PdfInspector.Inspect(pdf, OptionsFor(_recipient));

        Assert.Equal(Fixture.Title, info.Metadata.Title);
        Assert.Contains(Fixture.Text, PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, OptionsFor(_recipient)), StringComparison.Ordinal);
    }

    [Fact]
    public void AnyOfSeveralSuppliedCertificatesCanOpenTheDocument() {
        byte[] pdf = Fixture.Create(new Fixture.Settings(), _recipient);
        using X509Certificate2 unrelated = Fixture.CreateRecipient("OfficeIMO unrelated");
        var options = new PdfLoadOptions {
            PublicKeyDecryption = new PdfPublicKeyDecryptionOptions(OfficeSecurityProvider.Default, new[] { unrelated, _recipient })
        };

        Assert.Equal(Fixture.Title, PdfInspector.Inspect(pdf, options).Metadata.Title);
    }

    [Fact]
    public void MissingDecryptionOptionsRequireARecipientCertificate() {
        byte[] pdf = Fixture.Create(new Fixture.Settings(), _recipient);

        Assert.Throws<PdfRecipientCertificateRequiredException>(() => PdfReadDocument.Open(pdf));
        Assert.Throws<PdfRecipientCertificateRequiredException>(() => PdfReadDocument.Open(pdf, new PdfLoadOptions { Password = "anything" }));
    }

    [Fact]
    public void CertificateThatIsNotARecipientIsRejected() {
        byte[] pdf = Fixture.Create(new Fixture.Settings(), _recipient);

        Assert.Throws<PdfRecipientNotFoundException>(() => PdfReadDocument.Open(pdf, OptionsFor(_otherRecipient)));
    }

    [Fact]
    public void RecipientCertificateWithoutPrivateKeyReportsADecryptionFailure() {
        byte[] pdf = Fixture.Create(new Fixture.Settings(), _recipient);
#pragma warning disable SYSLIB0057 // Public certificate only, loaded the same way on net8.0 and net10.0.
        using var publicOnly = new X509Certificate2(_recipient.Export(X509ContentType.Cert));
#pragma warning restore SYSLIB0057

        PdfRecipientDecryptionException exception = Assert.Throws<PdfRecipientDecryptionException>(() =>
            PdfReadDocument.Open(pdf, OptionsFor(publicOnly)));
        Assert.Contains("EnvelopePrivateKeyMissing", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("adbe.pkcs7.s6", 5, null, null)]
    [InlineData("adbe.pkcs7.s4", 4, null, null)]
    [InlineData("adbe.pkcs7.s5", 2, null, null)]
    [InlineData("adbe.pkcs7.s5", 5, "AESV2", null)]
    [InlineData("adbe.pkcs7.s5", 5, "None", null)]
    [InlineData("adbe.pkcs7.s5", 5, null, "OtherCryptFilter")]
    public void UnsupportedOrInconsistentLayoutsFailClosed(string subFilter, int version, string? cfm, string? stringFilter) {
        byte[] pdf = Fixture.Create(new Fixture.Settings {
            SubFilter = subFilter,
            Version = version,
            CfmOverride = cfm,
            StringFilterOverride = stringFilter
        }, _recipient);

        Assert.Throws<PdfUnsupportedEncryptionException>(() => PdfReadDocument.Open(pdf, OptionsFor(_recipient)));
    }

    [Theory]
    [InlineData("40", "5")]
    [InlineData(null, "40")]
    [InlineData("40", null)]
    public void Rc4CryptFilterKeyLengthComesFromTheFilter(string? documentLength, string? filterLength) {
        byte[] pdf = Fixture.Create(new Fixture.Settings {
            Version = 4,
            Method = Fixture.Cipher.Rc4,
            KeyLengthBits = 40,
            DocumentLengthOverride = documentLength ?? "null",
            FilterLengthOverride = filterLength ?? "null"
        }, _recipient);

        Assert.Equal(Fixture.Title, PdfInspector.Inspect(pdf, OptionsFor(_recipient)).Metadata.Title);
    }

    [Fact]
    public void DisagreeingCryptFilterAndDocumentKeyLengthsFailClosed() {
        byte[] pdf = Fixture.Create(new Fixture.Settings {
            Version = 4,
            Method = Fixture.Cipher.Rc4,
            KeyLengthBits = 40,
            DocumentLengthOverride = "128"
        }, _recipient);

        Assert.Throws<PdfUnsupportedEncryptionException>(() => PdfReadDocument.Open(pdf, OptionsFor(_recipient)));
    }

    [Theory]
    [InlineData("DefaultCryptFilter", true)]
    [InlineData("Identity", false)]
    public void EmbeddedFileFilterMustMatchTheStreamFilter(string embeddedFileFilter, bool opens) {
        byte[] pdf = Fixture.Create(new Fixture.Settings { EmbeddedFileFilter = embeddedFileFilter }, _recipient);

        if (opens) {
            Assert.Equal(Fixture.Title, PdfInspector.Inspect(pdf, OptionsFor(_recipient)).Metadata.Title);
        } else {
            Assert.Throws<PdfUnsupportedEncryptionException>(() => PdfReadDocument.Open(pdf, OptionsFor(_recipient)));
        }
    }

    [Fact]
    public void RecipientDataWithUnexpectedLengthFailsClosed() {
        byte[] pdf = Fixture.Create(new Fixture.Settings { RecipientContentLength = 20 }, _recipient);

        Assert.Throws<PdfUnsupportedEncryptionException>(() => PdfReadDocument.Open(pdf, OptionsFor(_recipient)));
    }

    [Fact]
    public void RecipientPermissionsAreEnforced() {
        int permissions = Fixture.AllPermissions & ~(1 << 4) & ~(1 << 9);
        byte[] pdf = Fixture.Create(new Fixture.Settings { Permissions = permissions }, _recipient);

        Assert.Throws<PdfPermissionDeniedException>(() =>
            PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, OptionsFor(_recipient)));
        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, new PdfLoadOptions {
            PublicKeyDecryption = new PdfPublicKeyDecryptionOptions(OfficeSecurityProvider.Default, _recipient),
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        });
        Assert.Contains(Fixture.Text, text, StringComparison.Ordinal);
    }

    [Fact]
    public void RecipientsCannotRemoveEncryption() {
        byte[] pdf = Fixture.Create(new Fixture.Settings(), _recipient);

        Assert.Throws<PdfMutationBlockedException>(() => PdfSecurityEditor.Decrypt(pdf, "owner", OptionsFor(_recipient)));
    }

    [Fact]
    public void LoadOptionCopiesKeepPublicKeyDecryption() {
        PdfLoadOptions options = OptionsFor(_recipient);

        Assert.Same(options.PublicKeyDecryption, options.Clone().PublicKeyDecryption);
        Assert.Same(options.PublicKeyDecryption, options.Clone(1024, "password").PublicKeyDecryption);
    }

    [Fact]
    public void DecryptionOptionsValidateTheirInputs() {
        Assert.Throws<ArgumentNullException>(() => new PdfPublicKeyDecryptionOptions(null!, _recipient));
        Assert.Throws<ArgumentException>(() => new PdfPublicKeyDecryptionOptions(OfficeSecurityProvider.Default, Array.Empty<X509Certificate2>()));
        Assert.Throws<ArgumentException>(() => new PdfPublicKeyDecryptionOptions(OfficeSecurityProvider.Default, new X509Certificate2[] { null! }));
    }

    [Theory]
    [InlineData("pyhanko-s4-rc4-128.pdf", "adbe.pkcs7.s4", true)]
    [InlineData("pyhanko-s5-aesv2.pdf", "adbe.pkcs7.s5", true)]
    [InlineData("pyhanko-s5-aesv3.pdf", "adbe.pkcs7.s5", false)]
    [InlineData("pyhanko-s5-aesv3-plain-metadata.pdf", "adbe.pkcs7.s5", true)]
    public void OpensPublicKeyPdfsWrittenByPyHanko(string fileName, string subFilter, bool allowsPrinting) {
        byte[] pdf = File.ReadAllBytes(Path.Combine(PyHankoFixtureRoot, fileName));
        using X509Certificate2 recipient = PyHankoRecipient();
        PdfLoadOptions options = OptionsFor(recipient);

        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, options);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf, options);

        Assert.Contains("Public key fixture", text, StringComparison.Ordinal);
        Assert.Equal("Public key fixture title", info.Metadata.Title);
        Assert.Equal(subFilter, info.Security.EncryptionSubFilter);
        Assert.Equal(PdfPasswordAuthenticationRole.PublicKeyRecipient, info.Security.PasswordAuthenticationRole);
        Assert.Equal(allowsPrinting, info.Security.AllowsPrinting);
        Assert.Throws<PdfRecipientNotFoundException>(() => PdfReadDocument.Open(pdf, OptionsFor(_recipient)));
    }

    [Fact]
    public void PyHankoAesGcmPublicKeyPdfIsRefused() {
        byte[] pdf = File.ReadAllBytes(Path.Combine(PyHankoFixtureRoot, "pyhanko-s5-aes-gcm.pdf"));
        using X509Certificate2 recipient = PyHankoRecipient();

        Assert.Throws<PdfUnsupportedEncryptionException>(() => PdfReadDocument.Open(pdf, OptionsFor(recipient)));
    }

    private static PdfLoadOptions OptionsFor(X509Certificate2 certificate) => new() {
        PublicKeyDecryption = new PdfPublicKeyDecryptionOptions(OfficeSecurityProvider.Default, certificate)
    };

    private static X509Certificate2 PyHankoRecipient() =>
        Fixture.LoadPfx(Convert.FromBase64String(PyHankoRecipientPkcs12), "officeimo-test");

    // Throwaway RSA recipient written by generate_pyhanko_pubkey.py as recipient.pfx (password "officeimo-test").
    // It protects only the pyHanko fixtures; the repository does not track .pfx files, so it is kept here.
    private const string PyHankoRecipientPkcs12 =
        "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";

    private static string PyHankoFixtureRoot =>
        Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "PublicKeyEncryption");
}
