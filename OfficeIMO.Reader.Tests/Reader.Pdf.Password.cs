using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfPasswordTests {
    [Fact]
    public void PasswordDoesNotBypassExtractionPermissionsOrInputLimits() {
        byte[] plain = PdfDocument.Create(pdf => pdf.Page(page => page.Content(content => content.Text("Restricted text.")))).ToBytes();
        byte[] encrypted = PdfDocument.Load(plain).Security.Encrypt(new("reader") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.Accessibility
        }).Pdf;
        Assert.Throws<PdfInvalidPasswordException>(() => PdfReaderAdapter.ReadDocument(encrypted,
            pdfOptions: new ReaderPdfOptions { Password = "incorrect" }));
        Assert.Throws<PdfPermissionDeniedException>(() => PdfReaderAdapter.ReadDocument(encrypted,
            pdfOptions: new ReaderPdfOptions { Password = "reader" }));
        Assert.Throws<IOException>(() => PdfReaderAdapter.ReadDocument(encrypted,
            readerOptions: new ReaderOptions { MaxInputBytes = encrypted.Length - 1 },
            pdfOptions: new ReaderPdfOptions { Password = "owner" }));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void AuthorizedEncryptedPdfSupportsChunkAndDocumentPathAndStream(bool document, bool stream) {
        const string password = "local-reader-secret";
        byte[] plain = PdfDocument.Create(pdf => pdf.Page(page => page.Content(content => content.Text("Protected source text.")))).ToBytes();
        byte[] encrypted = PdfDocument.Load(plain).Security.Encrypt(new(password) {
            OwnerPassword = "owner-secret", AllowedPermissions = PdfStandardPermissions.CopyContents
        }).Pdf;
        string path = Path.Combine(Path.GetTempPath(), "officeimo-reader-password-" + Guid.NewGuid().ToString("N") + ".pdf");
        File.WriteAllBytes(path, encrypted);
        try {
            var options = new ReaderPdfOptions { Password = password }.Clone();
            using var input = new MemoryStream(encrypted, writable: false);
            string text = document
                ? System.Text.Json.JsonSerializer.Serialize(stream
                    ? PdfReaderAdapter.ReadDocument(input, "protected.pdf", pdfOptions: options)
                    : PdfReaderAdapter.ReadDocument(path, pdfOptions: options))
                : string.Join("\n", (stream
                    ? PdfReaderAdapter.Read(input, "protected.pdf", pdfOptions: options)
                    : PdfReaderAdapter.Read(path, pdfOptions: options)).Select(chunk => chunk.Text));
            Assert.Contains("Protected source text.", text);
            Assert.DoesNotContain(password, text);
            Assert.DoesNotContain("owner-secret", text);
            Assert.Equal(encrypted, File.ReadAllBytes(path));
        } finally { File.Delete(path); }
    }
}
