using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlManifestSchemesApplyUrlTabAndNewlinePreprocessing() {
        string html = "<html><head><link rel=\"c2pa-manifest\" href=\"java&#10;script:alert(1)\"></head></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ActiveKeyframeImageCarriersAreRemoved() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>@keyframes pulse{from{background-image:url('" + dataUri +
            "')}}.box{animation:pulse 1s}</style><div class=\"box\"></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void UnlayeredCustomPropertyOverridesLaterLayeredDeclaration() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.box{--img:url('" + dataUri +
            "');background-image:var(--img)}@layer theme{.box{--img:none}}</style><div class=\"box\"></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void EpubRootfilePathsDecodePercentEncodedSegments() {
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/container.xml", Encoding.UTF8.GetBytes(
                "<container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" version=\"1.0\"><rootfiles>" +
                "<rootfile full-path=\"OPS/package%20document.opf\" media-type=\"application/oebps-package+xml\"/>" +
                "</rootfiles></container>")),
            ("OPS/package document.opf", Encoding.UTF8.GetBytes(
                "<package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" unique-identifier=\"id\">" +
                "<metadata><identifier xmlns=\"http://purl.org/dc/elements/1.1/\" id=\"id\">fixture</identifier></metadata>" +
                "<manifest/><spine/></package>")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = EpubDocument.RemoveProvenance(package);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }
}
