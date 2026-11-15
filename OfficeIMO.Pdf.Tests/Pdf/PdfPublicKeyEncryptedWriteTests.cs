using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Security;
using Xunit;
using Fixture = OfficeIMO.Tests.Pdf.PublicKeyEncryptedPdfFixture;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPublicKeyEncryptedWriteTests {
    // RSA key generation dominates the run time, so the recipients are shared by every test.
    private static readonly X509Certificate2 _recipient = Fixture.CreateRecipient("OfficeIMO PDF writer recipient");
    private static readonly X509Certificate2 _secondRecipient = Fixture.CreateRecipient("OfficeIMO PDF second recipient");
    private static readonly X509Certificate2 _outsider = Fixture.CreateRecipient("OfficeIMO PDF outsider");
    private const string Text = "Certificate encrypted text";
    private const string Title = "Certificate encrypted title";

    private static byte[] CreateEncrypted(PdfStandardPermissions permissions, bool encryptMetadata = true) {
        PdfDocument source = PdfDocument.Load(PdfDocument.Create()
            .Meta(title: Title)
            .Paragraph(paragraph => paragraph.AddText(Text))
            .ToBytes());
        var encryption = new PdfPublicKeyEncryptionOptions(new OfficeSecurityProvider(), new[] { _recipient, _secondRecipient }) {
            AllowedPermissions = permissions,
            EncryptMetadata = encryptMetadata
        };
        return source.Security.EncryptForRecipients(encryption).Pdf;
    }

    private static PdfLoadOptions OptionsFor(X509Certificate2 certificate) =>
        new() { PublicKeyDecryption = new PdfPublicKeyDecryptionOptions(new OfficeSecurityProvider(), certificate) };

    [Fact]
    public void EveryRecipientOpensTheDocumentWithTheGrantedPermissions() {
        PdfStandardPermissions granted = PdfStandardPermissions.All & ~PdfStandardPermissions.Print & ~PdfStandardPermissions.HighQualityPrint;
        byte[] pdf = CreateEncrypted(granted);

        Assert.DoesNotContain(Text, PdfEncoding.Latin1GetString(pdf), StringComparison.Ordinal);
        foreach (X509Certificate2 certificate in new[] { _recipient, _secondRecipient }) {
            PdfDocumentInfo info = PdfInspector.Inspect(pdf, OptionsFor(certificate));
            Assert.Equal(Title, info.Metadata.Title);
            Assert.Equal("Adobe.PubSec", info.Security.EncryptionFilter);
            Assert.Equal("adbe.pkcs7.s5", info.Security.EncryptionSubFilter);
            Assert.Equal(granted, info.Security.AllowedStandardPermissions);
            Assert.False(info.Security.AllowsPrinting);
            Assert.Contains(Text, PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, OptionsFor(certificate)), StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RecipientsWithoutCopyPermissionStillGetTheDocument() {
        const PdfStandardPermissions granted = PdfStandardPermissions.Print;
        byte[] pdf = CreateEncrypted(granted);

        PdfDocumentInfo info = PdfInspector.Inspect(pdf, OptionsFor(_recipient));
        Assert.Equal(granted, info.Security.AllowedStandardPermissions);
        Assert.Throws<PdfPermissionDeniedException>(() =>
            PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, OptionsFor(_recipient)));
    }

    [Fact]
    public void OtherCertificatesAndMissingOptionsFailClosed() {
        byte[] pdf = CreateEncrypted(PdfStandardPermissions.All);

        Assert.Throws<PdfRecipientNotFoundException>(() => PdfInspector.Inspect(pdf, OptionsFor(_outsider)));
        Assert.Throws<PdfRecipientCertificateRequiredException>(() => PdfInspector.Inspect(pdf));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RecipientEntriesFollowIsoKeyDerivationIndependentlyOfOfficeIMO(bool encryptMetadata) {
        // Opens the output without OfficeIMO.Pdf: Bouncy Castle decrypts the CMS recipient entry, and the file key is
        // SHA-256 over the seed, the recipient entry and, with clear metadata, four 0xFF bytes (ISO 32000-2 7.6.5.3).
        const PdfStandardPermissions granted = PdfStandardPermissions.CopyContents | PdfStandardPermissions.Accessibility;
        byte[] pdf = CreateEncrypted(granted, encryptMetadata);
        string raw = PdfEncoding.Latin1GetString(pdf);

        Assert.Contains("/Filter /Adobe.PubSec /SubFilter /adbe.pkcs7.s5 /V 5", raw, StringComparison.Ordinal);
        Assert.Contains("/CFM /AESV3", raw, StringComparison.Ordinal);
        byte[] envelope = FromHex(Regex.Match(raw, @"/Recipients \[<([0-9A-Fa-f]+)>\]").Groups[1].Value);
        var enveloped = new CmsEnvelopedData(envelope);
        RecipientInformation recipient = Assert.Single(enveloped.GetRecipientInfos().GetRecipients().Cast<RecipientInformation>(),
            info => info.RecipientID.Match(DotNetUtilities.FromX509Certificate(_recipient)));
        byte[] content = recipient.GetContent(DotNetUtilities.GetRsaKeyPair(_recipient.GetRSAPrivateKey()!).Private);
        Assert.Equal(24, content.Length);
        int permissions = (content[20] << 24) | (content[21] << 16) | (content[22] << 8) | content[23];
        Assert.Equal(PdfStandardEncryptionOptions.ToRawPermissions(granted) | 1, permissions);

        byte[] fileKey;
        using (SHA256 sha = SHA256.Create()) {
            byte[] material = content.Take(20).Concat(envelope).Concat(encryptMetadata ? Array.Empty<byte>() : new byte[] { 0xFF, 0xFF, 0xFF, 0xFF }).ToArray();
            fileKey = sha.ComputeHash(material);
        }

        byte[] title = FromHex(Regex.Match(raw, @"/Title\s*<([0-9A-Fa-f]+)>").Groups[1].Value);
        using Aes aes = Aes.Create();
        aes.Key = fileKey;
        aes.IV = title.Take(16).ToArray();
        byte[] plain = aes.CreateDecryptor().TransformFinalBlock(title, 16, title.Length - 16);
        Assert.Equal(Title, PdfTextStringForTests(plain));
    }

    private static string PdfTextStringForTests(byte[] value) =>
        value.Length >= 2 && value[0] == 0xFE && value[1] == 0xFF
            ? System.Text.Encoding.BigEndianUnicode.GetString(value, 2, value.Length - 2)
            : PdfEncoding.Latin1GetString(value);

    private static byte[] FromHex(string hex) {
        var bytes = new byte[hex.Length / 2];
        for (int index = 0; index < bytes.Length; index++) bytes[index] = Convert.ToByte(hex.Substring(index * 2, 2), 16);
        return bytes;
    }
}
