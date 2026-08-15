using System.IO.Compression;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void SrcsetRejectsLeadingPlusDensityAndKeepsTheAcceptedSourceOffset() {
        const string url = "data:image/png;base64,AAAA";
        string srcset = url + " +1x, " + url + " 1x";

        HtmlSrcSetCandidate candidate = Assert.Single(HtmlSrcSetParser.Parse(srcset));

        Assert.Equal(url, candidate.Url);
        Assert.Equal(srcset.LastIndexOf(url, StringComparison.Ordinal), candidate.UrlStart);
    }

    [Fact]
    public void CssStateInsideNotDoesNotHideAReachableImageCarrier() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<html><head><style>.card:not(:hover){background-image:url('" + dataUri +
            "')}</style></head><body><div class=\"card\"></div></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void OpenDocumentRemovalPreservesUnrelatedContentTypesMetadata() {
        const string mediaType = "application/vnd.oasis.opendocument.text";
        const string manifestNamespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes(mediaType)),
            ("content.xml", Encoding.UTF8.GetBytes("<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>")),
            ("META-INF/manifest.xml", Encoding.UTF8.GetBytes(
                "<manifest:manifest xmlns:manifest=\"" + manifestNamespace + "\">" +
                "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"" + mediaType + "\"/>" +
                "<manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>" +
                "<manifest:file-entry manifest:full-path=\"META-INF/content_credential.c2pa\" manifest:media-type=\"application/c2pa\"/>" +
                "</manifest:manifest>")),
            ("[Content_Types].xml", Encoding.UTF8.GetBytes("not-opc-xml")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OdfDocument.RemoveProvenance(package);
        using var cleaned = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);

        Assert.Null(cleaned.GetEntry("META-INF/content_credential.c2pa"));
        Assert.Equal("not-opc-xml", ReadWave33Entry(cleaned, "[Content_Types].xml"));
    }

    private static byte[] CreateStoredPackage(params (string Name, byte[] Data)[] entries) {
        DateTimeOffset timestamp = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero);
        var outputEntries = entries.Select(entry => new OfficeProvenanceZipWriteEntry(
            entry.Name,
            entry.Data.Length,
            compress: false,
            timestamp,
            internalAttributes: 0,
            externalAttributes: 0,
            Array.Empty<byte>(),
            Array.Empty<byte>(),
            Array.Empty<byte>(),
            () => new MemoryStream(entry.Data, writable: false))).ToArray();
        return OfficeProvenanceZipWriter.Write(outputEntries, 1024 * 1024);
    }
}
