using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void DataUriMediaTypeUsesOnlyHtmlAsciiWhitespace() {
        string payload = Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<html><body><img src=\"data:\u00A0image/png;base64," + payload + "\"></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void PictureSourceTypeUsesOnlyHtmlAsciiWhitespace() {
        byte[] carrier = CreatePngWithManifest(CreateManifestStore());
        byte[] plainImage = OfficeProvenanceRemover.Remove(carrier, "image.png").ToArray();
        string source = Convert.ToBase64String(plainImage);
        string fallback = Convert.ToBase64String(carrier);
        string html = "<html><body><picture>" +
            "<source type=\"\u00A0image/png\" srcset=\"data:image/png;base64," + source + "\">" +
            "<img src=\"data:image/png;base64," + fallback + "\">" +
            "</picture></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Theory]
    [InlineData("div,,span")]
    [InlineData(",div")]
    [InlineData("div,")]
    public void InvalidCssSelectorListsAreInert(string selector) {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<html><head><style>" + selector + "{background-image:url('" + dataUri + "')}</style></head>" +
            "<body><div></div><span></span></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }
}
