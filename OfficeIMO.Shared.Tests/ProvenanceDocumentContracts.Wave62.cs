using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    private static string Wave62DataUri() => "data:image/png;name=\\\"fixture\\\";base64," +
        Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));

    [Fact]
    public void QuotedCssUrlsHonorEscapedClosingQuotes() {
        string html = "<style>.box{background-image:url(\"" + Wave62DataUri() + "\")}</style><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void BodyCustomPropertiesDoNotFlowIntoHeadElements() {
        string html = "<style>body{--img:url('" + Wave61DataUri() + "')}head{background-image:var(--img)}</style>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void BareStatefulSelectorsConservativelyResolveCustomPropertyImages() {
        string html = "<style>:root{--img:url('" + Wave61DataUri() + "')}:hover{background-image:var(--img)}</style><div></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
    }
}
