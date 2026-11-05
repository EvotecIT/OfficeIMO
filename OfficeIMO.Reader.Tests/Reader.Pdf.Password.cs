using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfPasswordTests {
    [Fact]
    public void PasswordDoesNotBypassExtractionPermissionsOrInputLimits() {
        byte[] plain = PdfDocument.Create(pdf => pdf.Page(page => page.Content(content => content.AddParagraph("Restricted text.")))).ToBytes();
        byte[] encrypted = PdfDocument.Load(plain).Security.Encrypt(new("reader") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.Accessibility
        }).Pdf;
        Assert.Throws<PdfInvalidPasswordException>(() => PdfReaderAdapter.ReadDocument(encrypted,
            pdfOptions: new ReaderPdfOptions { Password = "incorrect" }));
        Assert.Throws<PdfPermissionDeniedException>(() => PdfReaderAdapter.ReadDocument(encrypted,
            pdfOptions: new ReaderPdfOptions { Password = "reader" }));
        ReaderAdmissionLimitException error = Assert.Throws<ReaderAdmissionLimitException>(() => PdfReaderAdapter.ReadDocument(encrypted,
            readerOptions: new ReaderOptions { MaxInputBytes = encrypted.Length - 1 },
            pdfOptions: new ReaderPdfOptions { Password = "owner" }));
        Assert.Equal(PdfReadLimitKind.InputBytes, Assert.IsType<PdfReadLimitException>(error.InnerException).Kind);
        Assert.Equal(encrypted.Length - 1, error.Maximum);
        Assert.Equal(encrypted.Length, error.Actual);
    }

    [Theory]
    [InlineData(ReaderProjection.Full)]
    [InlineData(ReaderProjection.TextAndMetadata)]
    public void CommonPasswordAndNativePolicySurviveReaderProjection(ReaderProjection projection) {
        const string password = "common-reader-secret";
        byte[] plain = PdfDocument.Create(pdf => pdf.Page(page => page.Content(content => content.AddParagraph("Protected common text.")))).ToBytes();
        byte[] encrypted = PdfDocument.Load(plain).Security.Encrypt(new(password) {
            OwnerPassword = "owner-secret", AllowedPermissions = PdfStandardPermissions.Accessibility
        }).Pdf;
        var policy = new PdfLoadOptions { PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions };
        var format = new ReaderPdfOptions { LoadOptions = policy, ReadOptions = policy };
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(format).Build();
        var common = new ReaderOptions { OpenPassword = password, Projection = projection };
        OfficeDocumentReadResult result = reader.ReadDocument(encrypted, "protected.pdf", common);
        string text = System.Text.Json.JsonSerializer.Serialize(result);
        Assert.Contains("Protected common text.", text);
        Assert.DoesNotContain(password, text);
        Assert.DoesNotContain("owner-secret", text);
        Assert.Same(policy, format.LoadOptions);
        Assert.Same(policy, format.ReadOptions);
        Assert.Equal(PdfPermissionPolicy.IgnoreRestrictions, policy.PermissionPolicy);

        Assert.Throws<PdfInvalidPasswordException>(() => reader.ReadDocument(encrypted, "protected.pdf",
            new ReaderOptions { OpenPassword = "incorrect", Projection = projection }));
        var enforced = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();
        Assert.Throws<PdfPermissionDeniedException>(() => enforced.ReadDocument(encrypted, "protected.pdf", common));
        IOException error = Assert.Throws<IOException>(() => reader.ReadDocument(encrypted, "protected.pdf",
            new ReaderOptions { OpenPassword = password, Projection = projection, MaxInputBytes = encrypted.Length - 1 }));
        Assert.Contains("MaxInputBytes", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(password, error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void AuthorizedEncryptedPdfSupportsChunkAndDocumentPathAndStream(bool document, bool stream) {
        const string password = "local-reader-secret";
        byte[] plain = PdfDocument.Create(pdf => pdf.Page(page => page.Content(content => content.AddParagraph("Protected source text.")))).ToBytes();
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
