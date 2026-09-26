using System.IO;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlPreflightBoundsUnmatchedEndTagStackSearches() {
        string html = string.Concat(Enumerable.Repeat("<a>", 2000)) +
            string.Concat(Enumerable.Repeat("</b>", 2000));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            HtmlProvenance.Inspect(html, new OfficeProvenanceOptions { MaxContainerEntries = 3000 }));

        Assert.Contains("preflight scan limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlPreflightCountsTableElementsAfterSelectBreakout() {
        string html = "<table><select>" +
            string.Concat(Enumerable.Repeat("<tr><td><div></div></td></tr>", 20)) +
            "</table>";

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            HtmlProvenance.Inspect(html, new OfficeProvenanceOptions { MaxContainerEntries = 8 }));

        Assert.Contains("container-entry limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlRemovalScrubsDormantCssAndResolvedVarFallbacks() {
        string image = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style media='print'>.unmatched{--hero/**/:none;background-image:" +
            "var(--hero,url('" + image + "'))}</style><div class='other'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain(image, Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRemovalScrubsCommentSeparatedCustomPropertyUrl() {
        string image = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.box{--hero/**/:url('" + image + "');background:var(--hero)}</style>" +
            "<div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain(image, Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRemovalScrubsQuotedCssUrlWithLineContinuation() {
        string image = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        int split = image.IndexOf(',') + 16;
        string escaped = image.Substring(0, split) + "\\\n" + image.Substring(split);
        string html = "<style>.box{background-image:url(\"" + escaped + "\")}</style><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void HtmlRemovalScrubsDirectPictureSourceAndFallback() {
        string image = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<picture><source src='" + image + "'><img src='" + image + "'></picture>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void LegacyHtmlEncodingUsesCharacterReferencesWithoutAnExpandedString() {
        Encoding ascii = OfficeCharacterReferenceEncoding.WithCharacterReferenceFallback(Encoding.ASCII);

        Assert.Equal("A&#xFFFD;&#x1F600;B", ascii.GetString(ascii.GetBytes("A\uFFFD\U0001F600B")));
    }

    [Fact]
    public void LegacyHtmlRemovalChecksExpandedOutputBeforeWriting() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string html = "<html><head><meta charset='windows-1252'><link rel='c2pa-manifest' " +
            "href='claim.c2pa'></head><body>" + string.Concat(Enumerable.Repeat("&#x1F600;", 2000)) +
            "</body></html>";
        string input = Path.Combine(Path.GetTempPath(), $"OfficeIMO-html-budget-{Guid.NewGuid():N}.html");
        string output = Path.Combine(Path.GetTempPath(), $"OfficeIMO-html-budget-{Guid.NewGuid():N}.out.html");
        try {
            File.WriteAllBytes(input, windows1252.GetBytes(html));
            var options = new OfficeProvenanceRemovalOptions { MaxOutputBytes = 4096 };

            Assert.Throws<InvalidDataException>(() => HtmlProvenance.RemoveFile(input, output, options));
            Assert.False(File.Exists(output));
        } finally {
            if (File.Exists(input)) File.Delete(input);
            if (File.Exists(output)) File.Delete(output);
        }
    }
}
