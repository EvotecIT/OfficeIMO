using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailMimeProtectionTests {
    [Theory]
    [InlineData("application/pkcs7-signature", EmailProtectionKind.SmimeClearSigned)]
    [InlineData("application/pgp-signature", EmailProtectionKind.PgpMimeClearSigned)]
    public void DetectsAndPassesThroughUnchangedSignedMime(string protocol, EmailProtectionKind expectedKind) {
        byte[] source = CreateSignedMessage(protocol);

        EmailReadResult read = new EmailDocumentReader().Read(source);
        byte[] rewritten = new EmailDocumentWriter().ToBytes(read.Document, EmailFileFormat.Eml);

        Assert.Equal(expectedKind, read.Document.Protection.Kind);
        Assert.NotNull(read.Document.RawSource);
        Assert.Equal(source, rewritten);
    }

    [Fact]
    public void BlocksSignedMimeAfterTheModelChanges() {
        EmailDocument document = new EmailDocumentReader().Read(
            CreateSignedMessage("application/pkcs7-signature")).Document;
        document.Subject = "Changed";
        var writer = new EmailDocumentWriter();

        EmailConversionReport report = writer.AnalyzeConversion(document, EmailFileFormat.Eml);
        using var destination = new MemoryStream();
        EmailWriteResult result = writer.Write(document, destination, EmailFileFormat.Eml);

        Assert.True(report.HasPotentialDataLoss);
        Assert.False(report.CanWrite);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_PROTECTED_CONTENT_REWRITE");
        Assert.True(result.HasErrors);
        Assert.Equal(0, destination.Length);
        Assert.Throws<InvalidDataException>(() => writer.ToBytes(document, EmailFileFormat.Eml));
    }

    [Fact]
    public void ExplicitWarnPolicyAllowsProtectedConversionAndReportsIt() {
        EmailDocument document = new EmailDocumentReader().Read(
            CreateSignedMessage("application/pkcs7-signature")).Document;
        document.Subject = "Changed";
        var writer = new EmailDocumentWriter(new EmailWriterOptions(
            conversionLossPolicy: EmailConversionLossPolicy.Warn));

        byte[] output = writer.ToBytes(document, EmailFileFormat.Eml, out EmailWriteResult result);

        Assert.NotEmpty(output);
        Assert.False(result.HasErrors);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_PROTECTED_CONTENT_REWRITE" &&
            diagnostic.Severity == EmailDiagnosticSeverity.Warning);
    }

    private static byte[] CreateSignedMessage(string protocol) {
        string source = "From: sender@example.com\r\n" +
            "To: recipient@example.com\r\n" +
            "Subject: Signed\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/signed; protocol=\"" + protocol + "\"; boundary=\"sig\"\r\n\r\n" +
            "--sig\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nHello\r\n" +
            "--sig\r\nContent-Type: " + protocol + "\r\nContent-Transfer-Encoding: base64\r\n\r\nAQID\r\n" +
            "--sig--\r\n";
        return Encoding.ASCII.GetBytes(source);
    }
}