using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReadStreamTests {
    [Fact]
    public void PdfReadDocument_Load_ReadsFromCurrentStreamPosition() {
        byte[] pdf = BuildPdf();
        using var stream = BuildPrefixedStream(pdf);
        stream.Position = 5;

        PdfReadDocument document = PdfReadDocument.Open(stream);

        Assert.Single(document.Pages);
        Assert.Equal("Stream read", document.Metadata.Title);
        Assert.Contains("Stream readable text", document.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_ExtractAllText_ReadsFromCurrentStreamPosition() {
        byte[] pdf = BuildPdf();
        using var stream = BuildPrefixedStream(pdf);
        stream.Position = 5;

        string text = PdfTextExtractor.ExtractAllText(stream);

        Assert.Contains("Stream readable text", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextExtractor_GetMetadata_ReadsFromPathAndStream() {
        byte[] pdf = BuildPdf();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-read-stream-" + Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllBytes(path, pdf);
            using var stream = BuildPrefixedStream(pdf);
            stream.Position = 5;

            var fromPath = PdfTextExtractor.GetMetadata(path);
            var fromStream = PdfTextExtractor.GetMetadata(stream);

            Assert.Equal("Stream read", fromPath.Title);
            Assert.Equal("OfficeIMO", fromPath.Author);
            Assert.Equal(fromPath.Title, fromStream.Title);
            Assert.Equal(fromPath.Author, fromStream.Author);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void ReadStreamApis_RejectNullAndUnreadableStreams() {
        Assert.Throws<ArgumentNullException>(() => PdfReadDocument.Open((Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText((Stream)null!));
        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.GetMetadata((Stream)null!));

        using var unreadable = new WriteOnlyStream();

        Assert.Throws<ArgumentException>(() => PdfReadDocument.Open(unreadable));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllText(unreadable));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.GetMetadata(unreadable));
    }

    [Fact]
    public void ReadPathApis_RejectNullAndWhitespacePaths() {
        Assert.Throws<ArgumentNullException>(() => PdfReadDocument.Open((string)null!));
        Assert.Throws<ArgumentException>(() => PdfReadDocument.Open(" "));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractAllText((string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractAllText(" "));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.ExtractTextByPage((string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.ExtractTextByPage(" "));

        Assert.Throws<ArgumentNullException>(() => PdfTextExtractor.GetMetadata((string)null!));
        Assert.Throws<ArgumentException>(() => PdfTextExtractor.GetMetadata(" "));

        Assert.Throws<ArgumentNullException>(() => PdfImageExtractor.ExtractImages((string)null!));
        Assert.Throws<ArgumentException>(() => PdfImageExtractor.ExtractImages(" "));
    }

    [Fact]
    public void ReadApis_RejectEncryptedPdfsWithClearUnsupportedDiagnostic() {
        byte[] encrypted = BuildEncryptedPdfMarker();

        AssertPreciseEncryptedReadFailure(() => PdfReadDocument.Open(encrypted));
        AssertPreciseEncryptedReadFailure(() => PdfTextExtractor.ExtractAllText(encrypted));
        AssertEncrypted(() => PdfTextExtractor.GetMetadata(encrypted));
        AssertPreciseEncryptedReadFailure(() => PdfImageExtractor.ExtractImages(encrypted));
        AssertEncrypted(() => PdfPageExtractor.ExtractPages(encrypted, 1));

        static void AssertPreciseEncryptedReadFailure(Action action) {
            var exception = Assert.Throws<PdfPasswordRequiredException>(action);
            Assert.Contains("Encrypted PDF requires a password.", exception.Message, StringComparison.Ordinal);
        }

        static void AssertEncrypted(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.True(
                exception is PdfEncryptionException ||
                exception.Message.Contains("Encrypted PDF", StringComparison.OrdinalIgnoreCase));
        }
    }

    [Fact]
    public void RewriteApis_RejectSignedPdfsWithClearUnsupportedDiagnostic() {
        byte[] signed = BuildSignedPdfMarker();

        AssertSigned(() => PdfPageExtractor.ExtractPages(signed, 1));
        AssertSigned(() => PdfPageExtractor.SplitPages(signed));
        AssertSigned(() => PdfPageEditor.DeletePages(signed, 1));
        AssertSigned(() => PdfMetadataEditor.UpdateMetadata(signed, title: "Updated"));
        AssertSigned(() => PdfMerger.Merge(signed));
        AssertSigned(() => PdfStamper.StampText(signed, "STAMP"));

        static void AssertSigned(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("Signed PDF files are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RewriteApis_RouteFormPdfsByOperationCapability() {
        byte[] form = PdfDocument.Create().TextField("Name", value: "OfficeIMO").ToBytes();

        AssertFormBlocked(() => PdfPageExtractor.ExtractPages(form, 1));
        AssertFormBlocked(() => PdfPageExtractor.SplitPages(form));
        AssertFormBlocked(() => PdfPageEditor.DeletePages(form, 1));
        AssertFormBlocked(() => PdfMetadataEditor.UpdateMetadata(form, title: "Updated"));
        AssertFormPreserved(PdfMerger.Merge(form));
        AssertFormBlocked(() => PdfStamper.StampText(form, "STAMP"));

        static void AssertFormPreserved(byte[] rewritten) => Assert.Single(PdfInspector.Inspect(rewritten).FormFields);
        static void AssertFormBlocked(Action action) {
            var exception = Assert.ThrowsAny<NotSupportedException>(action);
            Assert.Contains("PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.", exception.Message, StringComparison.Ordinal);
        }
    }


}
