using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderMhtmlModularTests {
    [Fact]
    public void HtmlHandlerRegistersMhtmlExtensions() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHtmlHandler().Build();

        ReaderHandlerCapability capability = Assert.Single(reader.GetCapabilities(), item =>
            item.Id == OfficeDocumentReaderBuilderHtmlExtensions.HandlerId);

        Assert.Contains(".mht", capability.Extensions);
        Assert.Contains(".mhtml", capability.Extensions);
    }

    [Fact]
    public void ReaderHtmlProjectsMhtmlTextAndEmbeddedAssets() {
        byte[] archive = CreateArchive();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHtmlHandler().Build();
        using var chunkStream = new MemoryStream(archive, writable: false);

        ReaderChunk[] chunks = reader.Read(chunkStream, "saved.mhtml").ToArray();

        Assert.Contains(chunks, chunk => chunk.Text.Contains("Saved archive", StringComparison.Ordinal));

        using var documentStream = new MemoryStream(archive, writable: false);
        OfficeDocumentReadResult result = reader.ReadDocument(documentStream, "saved.mhtml");

        Assert.Equal(ReaderInputKind.Html, result.Kind);
        Assert.Equal("MHTML document", result.Source.Title);
        Assert.Contains("officeimo.html.mhtml", result.CapabilitiesUsed);
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
    public void ReaderHtmlMhtmlHonorsReaderInputLimit() {
        byte[] archive = CreateArchive();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHtmlHandler().Build();
        using var stream = new MemoryStream(archive, writable: false);

        Exception exception = Assert.ThrowsAny<Exception>(() => reader.ReadDocument(stream, "saved.mht",
            new ReaderOptions { MaxInputBytes = 32 }));

        Assert.Contains("MaxInputBytes", exception.Message, StringComparison.OrdinalIgnoreCase);
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
