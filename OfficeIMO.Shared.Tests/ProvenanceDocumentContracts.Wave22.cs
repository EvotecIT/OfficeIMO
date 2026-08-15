using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlNormalizesDirectDataUriWhitespaceAndPreservesItsSourceRange() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><head></head><body><img src=\"  {dataUri}  \"></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.Single(report.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("src=\"  data:image/png;base64,", output, StringComparison.Ordinal);
        Assert.Contains("  \"", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlCommentPreflightRecognizesParseErrorEndBangTerminator() {
        string elements = string.Concat(Enumerable.Repeat("<div></div>", 16));
        string html = "<html><head><!--x--!>" + elements + "</head><body></body></html>";

        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Inspect(
            html,
            new OfficeProvenanceOptions { MaxContainerEntries = 8 }));
    }

    [Fact]
    public void HtmlIgnoresImageDataUrisInsideInertStyleElements() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><head><style type=\"text/plain\">.sample{{background:url({dataUri})}}</style></head><body></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(Encoding.UTF8.GetBytes(html), result.ToArray());
    }
}
