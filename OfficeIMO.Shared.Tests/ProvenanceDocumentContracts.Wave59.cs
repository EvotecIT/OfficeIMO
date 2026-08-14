using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void EpubRemovalIgnoresUnrelatedMalformedOpcMetadata() {
        const string container =
            "<container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" version=\"1.0\"><rootfiles>" +
            "<rootfile full-path=\"OPS/package.opf\" media-type=\"application/oebps-package+xml\"/>" +
            "</rootfiles></container>";
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/container.xml", Encoding.UTF8.GetBytes(container)),
            ("OPS/package.opf", Encoding.UTF8.GetBytes(
                "<package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" unique-identifier=\"id\">" +
                "<metadata><identifier xmlns=\"http://purl.org/dc/elements/1.1/\" id=\"id\">fixture</identifier></metadata>" +
                "<manifest/><spine/></package>")),
            ("[Content_Types].xml", Encoding.UTF8.GetBytes("not OPC XML")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = EpubDocument.RemoveProvenance(package);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void CssAnimationNamesRemainCaseSensitive() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>@keyframes Pulse{from{background-image:url('" + dataUri +
            "')}}.box{animation:pulse 1s}</style><div class=\"box\"></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void InvalidUnquotedCssUrlTokensAreIgnored() {
        string dataUri = "data:image/svg+xml," + Uri.EscapeDataString(
            "<svg xmlns=\"http://www.w3.org/2000/svg\"><metadata><x:xmpmeta xmlns:x=\"adobe:ns:meta/\"/></metadata></svg>");
        string html = "<style>.box{background-image:url(" + dataUri + "\tinvalid)}</style><div class=\"box\"></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void VisioIgnoresOrphanConventionalApplicationMetadata() {
        byte[] package = CreateVisioPackageWithApplicationSignatureOnly(relationshipOwned: false);

        OfficeProvenanceRemovalResult result = VisioDocument.RemoveProvenance(package, "drawing.vsdx");

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }
}
