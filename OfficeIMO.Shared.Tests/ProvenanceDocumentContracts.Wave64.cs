using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void OdfRejectsBackslashEntryNames() {
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.text")),
            ("META-INF/manifest.xml", Encoding.UTF8.GetBytes(ValidOdfManifestXml)),
            ("content.xml", Encoding.UTF8.GetBytes("<content/>")),
            ("media\\foreign.png", Array.Empty<byte>()),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        Assert.Throws<InvalidDataException>(() => OdfDocument.RemoveProvenance(package));
    }

    [Fact]
    public void ImagePreloadSourceSetReplacesHref() {
        string provenance = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        byte[] cleanPng = OfficeProvenanceRemover.Remove(CreatePngWithManifest(CreateManifestStore()), "fixture.png").ToArray();
        string clean = "data:image/png;base64," + Convert.ToBase64String(cleanPng);
        string html = "<link rel='preload' as='image' href='" + provenance + "' imagesrcset='" + clean + " 1x'>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Theory]
    [InlineData("object", "data")]
    [InlineData("embed", "src")]
    public void ObjectAndEmbedImageDataUrisAreProcessed(string element, string attribute) {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<" + element + " type='image/png' " + attribute + "='" + dataUri + "'></" + element + ">";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void EpubRejectsMalformedPercentEscapesInRootfilePaths() {
        string container = Wave63ContainerPrefix +
            "<rootfile full-path='OPS/%ZZ.opf' media-type='application/oebps-package+xml'/>" +
            Wave63ContainerSuffix;
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/container.xml", Encoding.UTF8.GetBytes(container)),
            ("OPS/%ZZ.opf", Encoding.UTF8.GetBytes(Wave63Opf)),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        Assert.Throws<InvalidDataException>(() => EpubDocument.RemoveProvenance(package));
    }
}
