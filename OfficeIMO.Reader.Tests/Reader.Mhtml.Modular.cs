using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Html;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderMhtmlModularTests {
    [Fact]
    public void HtmlHandlerDoesNotClaimMhtmlExtensions() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHtmlHandler().Build();

        ReaderHandlerCapability capability = Assert.Single(reader.GetCapabilities(), item =>
            item.Id == OfficeDocumentReaderBuilderHtmlExtensions.HandlerId);

        Assert.DoesNotContain(".mht", capability.Extensions);
        Assert.DoesNotContain(".mhtml", capability.Extensions);
    }

    [Fact]
    public void EmailPackageRegistersDedicatedMhtmlHandler() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddMhtmlHandler().Build();

        ReaderHandlerCapability capability = Assert.Single(reader.GetCapabilities(), item =>
            item.Id == OfficeDocumentReaderBuilderMhtmlExtensions.HandlerId);

        Assert.Contains(".mht", capability.Extensions);
        Assert.Contains(".mhtml", capability.Extensions);
        Assert.Equal(
            OfficeDocumentReaderBuilderMhtmlExtensions.DefaultMaxInputBytes,
            reader.GetHandlerDefaultMaxInputBytes("message.mhtml"));
    }

    [Fact]
    public void AggregateEmailRegistrationIncludesMhtml() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddEmailHandlers().Build();

        Assert.Contains(reader.GetCapabilities(), item =>
            item.Id == OfficeDocumentReaderBuilderMhtmlExtensions.HandlerId &&
            item.Extensions.Contains(".mhtml", StringComparer.OrdinalIgnoreCase));
    }

    [Fact]
    public void ReaderEmailProjectsMhtmlTextAndEmbeddedAssets() {
        byte[] archive = CreateArchive();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddMhtmlHandler().Build();
        using var chunkStream = new MemoryStream(archive, writable: false);

        ReaderChunk[] chunks = reader.Read(chunkStream, "saved.mhtml").ToArray();

        Assert.Contains(chunks, chunk => chunk.Text.Contains("Saved archive", StringComparison.Ordinal));

        using var documentStream = new MemoryStream(archive, writable: false);
        OfficeDocumentReadResult result = reader.ReadDocument(documentStream, "saved.mhtml");

        Assert.Equal(ReaderInputKind.Html, result.Kind);
        Assert.Equal("MHTML document", result.Source.Title);
        Assert.Contains("officeimo.reader.mhtml", result.CapabilitiesUsed);
        Assert.Contains("officeimo.mhtml", result.CapabilitiesUsed);
        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal("image", asset.Kind);
        Assert.Equal("image/png", asset.MediaType);
        Assert.Equal("cid:logo", asset.SourceObjectId);
        Assert.Equal(new byte[] { 1, 2, 3 }, asset.PayloadBytes);
        Assert.False(string.IsNullOrWhiteSpace(asset.PayloadHash));
        ReaderVisual visual = Assert.Single(result.Visuals);
        Assert.Equal(asset.PayloadHash, visual.PayloadHash);
        Assert.Equal("image/png", visual.MimeType);
    }

    [Fact]
    public void ReaderEmailMhtmlHonorsReaderInputLimit() {
        byte[] archive = CreateArchive();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddMhtmlHandler().Build();
        using var stream = new MemoryStream(archive, writable: false);

        Exception exception = Assert.ThrowsAny<Exception>(() => reader.ReadDocument(stream, "saved.mht",
            new ReaderOptions { MaxInputBytes = 32 }));

        Assert.Contains("MaxInputBytes", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ContentDetectedMhtmlUsesHandlerDefaultInputLimit() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-mhtml-limit-" + Guid.NewGuid().ToString("N"));
        try {
            byte[] header = Encoding.ASCII.GetBytes("<html><body>content-detected archive</body></html>");
            byte[] detectionPrefix = Enumerable.Repeat((byte)' ', 64 * 1024).ToArray();
            Buffer.BlockCopy(header, 0, detectionPrefix, 0, header.Length);
            using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                output.Write(detectionPrefix, 0, detectionPrefix.Length);
                output.SetLength(OfficeDocumentReaderBuilderMhtmlExtensions.DefaultMaxInputBytes + 1L);
            }

            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddMhtmlHandler().Build();

            Exception exception = Assert.ThrowsAny<Exception>(() => reader.ReadDocument(path));

            Assert.Contains("MaxInputBytes", exception.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static byte[] CreateArchive() {
        const string archive = "MIME-Version: 1.0\r\n" +
            "Subject: MHTML document\r\n" +
            "Content-Type: multipart/related; boundary=archive; type=\"text/html\"; start=\"<root>\"\r\n\r\n" +
            "--archive\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <root>\r\n" +
            "Content-Location: https://example.test/page/index.html\r\n\r\n" +
            "<html><body><h1>Saved archive</h1><img alt=\"Logo\" src=\"cid:logo\"></body></html>\r\n" +
            "--archive\r\nContent-Type: image/png\r\nContent-ID: <logo>\r\n" +
            "Content-Transfer-Encoding: base64\r\n\r\nAQID\r\n" +
            "--archive--\r\n";
        return Encoding.ASCII.GetBytes(archive);
    }
}
