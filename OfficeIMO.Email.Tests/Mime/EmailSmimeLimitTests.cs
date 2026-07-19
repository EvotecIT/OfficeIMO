using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailSmimeLimitTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RawEmlCmsPayloadHonorsConfiguredEncodedLimit(bool clearSigned) {
        byte[] cms = Enumerable.Range(0, 64).Select(value => (byte)value).ToArray();
        byte[] message = clearSigned
            ? CreateClearSignedMessage(cms)
            : CreateOpaqueMessage(cms);
        using EmailReadResult read = new EmailDocumentReader().Read(message);
        var options = new CmsVerificationOptions { MaxEncodedBytes = 8 };

        EmailSmimeVerificationResult result = EmailSmime.Verify(read.Document, options);

        Assert.Equal(clearSigned
            ? EmailProtectionKind.SmimeClearSigned
            : EmailProtectionKind.SmimeOpaque, result.ProtectionKind);
        Assert.Null(result.Cryptography);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_SMIME_PAYLOAD_UNAVAILABLE" &&
            diagnostic.Message.Contains("configured CMS limit", StringComparison.Ordinal));
    }

    [Fact]
    public void RawEmlInvalidBase64FallbackHonorsConfiguredEncodedLimit() {
        byte[] message = CreateOpaqueMessage(new string('!', 64));
        using EmailReadResult read = new EmailDocumentReader().Read(message);
        var options = new CmsVerificationOptions { MaxEncodedBytes = 8 };

        EmailSmimeVerificationResult result = EmailSmime.Verify(read.Document, options);

        Assert.Null(result.Cryptography);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_MIME_BASE64_INVALID");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_SMIME_PAYLOAD_UNAVAILABLE" &&
            diagnostic.Message.Contains("configured CMS limit", StringComparison.Ordinal));
    }

    private static byte[] CreateOpaqueMessage(byte[] cms) =>
        CreateOpaqueMessage(Convert.ToBase64String(cms));

    private static byte[] CreateOpaqueMessage(string transferPayload) => Encoding.ASCII.GetBytes(
        "MIME-Version: 1.0\r\n" +
        "Content-Type: application/pkcs7-mime; smime-type=signed-data\r\n" +
        "Content-Transfer-Encoding: base64\r\n\r\n" +
        transferPayload + "\r\n");

    private static byte[] CreateClearSignedMessage(byte[] signature) => Encoding.ASCII.GetBytes(
        "MIME-Version: 1.0\r\n" +
        "Content-Type: multipart/signed; protocol=\"application/pkcs7-signature\"; " +
        "boundary=\"officeimo-limit\"\r\n\r\n" +
        "--officeimo-limit\r\n" +
        "Content-Type: text/plain\r\n\r\n" +
        "Signed body\r\n" +
        "--officeimo-limit\r\n" +
        "Content-Type: application/pkcs7-signature\r\n" +
        "Content-Transfer-Encoding: base64\r\n\r\n" +
        Convert.ToBase64String(signature) + "\r\n" +
        "--officeimo-limit--\r\n");
}
