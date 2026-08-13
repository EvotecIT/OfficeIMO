using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlBogusCommentsEndAtTheFirstGreaterThanSign() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<html><head><script type=\"application/c2pa\">" + manifest +
            "</script></head><body><?x \"><div></div>" + string.Concat(Enumerable.Repeat("<span></span>", 32)) + "</body></html>";

        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Inspect(
            html, new OfficeProvenanceOptions { MaxContainerEntries = 16 }));
    }

    [Fact]
    public void HtmlPreflightPreservesForeignContentCdataAsText() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<html><head><script type=\"application/c2pa\">" + manifest +
            "</script></head><body><svg><![CDATA[" + string.Concat(Enumerable.Repeat("<div></div>", 64)) +
            "]]></svg></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html, new OfficeProvenanceOptions { MaxContainerEntries = 16 });

        Assert.Single(report.Evidence);
    }

    [Fact]
    public void HtmlNormalizesEmbeddedSvgDeclarationAfterUtf16Bom() {
        string svg = "<?xml version=\"1.0\" encoding=\"utf-16\"?><svg xmlns=\"http://www.w3.org/2000/svg\"><metadata><x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\"><rdf:Description xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\" iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta></metadata></svg>";
        byte[] encoded = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(svg)).ToArray();
        string dataUri = "data:image/svg+xml;charset=utf-16;base64," + Convert.ToBase64String(encoded);
        string html = $"<html><body><img src=\"{dataUri}\"></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void HtmlIgnoresImageDataUrisUsedAsMediaSources() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><body><video><source src=\"{dataUri}\"></video></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlStillProcessesInlineCssOnInactivePictureFallbackImages() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><body><picture><source media=\"screen\" srcset=\"image.png\"><img src=\"fallback.png\" style=\"background-image:url({dataUri})\"></picture></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void OdfRemovalCleansTheNativeManifestFileEntry() {
        const string manifestPath = "META-INF/content_credential.c2pa";
        string manifestXml =
            "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">" +
            "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/>" +
            $"<manifest:file-entry manifest:full-path=\"{manifestPath}\" manifest:media-type=\"application/c2pa\"/>" +
            "</manifest:manifest>";
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, manifestPath, CreateManifestStore(), CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/manifest.xml", manifestXml, CompressionLevel.Optimal);
            }
            package = RewriteFixtureWithStoredMimetype(output.ToArray());
        }

        OfficeProvenanceRemovalResult result = OdfDocument.RemoveProvenance(package, "document.odt");
        XDocument manifest = XDocument.Parse(Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "META-INF/manifest.xml")));

        Assert.DoesNotContain(manifest.Descendants(), element =>
            string.Equals((string?)element.Attribute(XName.Get("full-path", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")), manifestPath, StringComparison.Ordinal));
    }
}
