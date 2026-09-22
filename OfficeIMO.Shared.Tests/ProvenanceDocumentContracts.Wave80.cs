using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void SvgDataUriRewritePreservesQuotedMetadataWithSemicolon() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:c2pa='http://c2pa.org/manifest'>" +
            "<metadata><c2pa:manifest>" + manifest + "</c2pa:manifest></metadata></svg>";
        string dataUri = "data:image/svg+xml;profile=\"a;b\";charset=utf-8;base64," +
            Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove("<img src='" + dataUri + "'>");

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("profile=&quot;a;b&quot;", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void CssDataUriRewritePreservesQuotedMetadataWithSemicolon() {
        string image = "data:image/png;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string css = ".box{background-image:url('" + image + "')}";
        string stylesheet = "data:text/css;profile=\"a;b\";charset=utf-8;base64," +
            Convert.ToBase64String(Encoding.UTF8.GetBytes(css));
        string html = "<link rel='stylesheet' href='" + stylesheet + "'><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("profile=&quot;a;b&quot;", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void CssRewriteEscapesQuotesInsideDoubleQuotedDataUri() {
        string image = "data:image/png;profile=\\\"a;b\\\";base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.box{background-image:url(\"" + image + "\")}</style><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("profile=\\\"a;b\\\"", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }
}
