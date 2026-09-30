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

    [Theory]
    [MemberData(nameof(HtmlPreflightRejectsEntryExpansionAcrossLexicalStateTransitionsCases))]
    public void HtmlPreflightRejectsEntryExpansionAcrossLexicalStateTransitions(string caseName, string html, int maximumEntries) {
        _ = caseName;
        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Inspect(
            html, new OfficeProvenanceOptions { MaxContainerEntries = maximumEntries }));
    }

    public static IEnumerable<object[]> HtmlPreflightRejectsEntryExpansionAcrossLexicalStateTransitionsCases() {
        {
            string elements = string.Concat(Enumerable.Repeat("<div></div>", 16));
            string html = "<html><head><!--x--!>" + elements + "</head><body></body></html>";
            yield return new object[] { "HtmlCommentPreflightRecognizesParseErrorEndBangTerminator", html, 8 };
        }
        {
            string manifest = Convert.ToBase64String(CreateManifestStore());
            string html = "<html><head><script type=\"application/c2pa\">" + manifest +
                "</script></head><body><?x \"><div></div>" + string.Concat(Enumerable.Repeat("<span></span>", 32)) + "</body></html>";
            yield return new object[] { "HtmlBogusCommentsEndAtTheFirstGreaterThanSign", html, 16 };
        }
        {
            string html = "<html><body><svg><foreignObject><![CDATA[x>" +
                string.Concat(Enumerable.Repeat("<span></span>", 32)) +
                "</foreignObject></svg></body></html>";
            yield return new object[] { "HtmlPreflightTreatsForeignObjectChildrenAsHtml", html, 16 };
        }
        {
            string html = "<html><body><svg><p><![CDATA[hidden>" + string.Concat(Enumerable.Repeat("<div></div>", 32));
            yield return new object[] { "HtmlForeignContentBreakoutTagsRestoreHtmlTokenizationDuringPreflight", html, 12 };
        }
        {
            string html = "<html><body><div data-value=unquoted\">" +
                string.Concat(Enumerable.Repeat("<span></span>", 16)) +
                "</div></body></html>";
            yield return new object[] { "HtmlPreflightTreatsQuotesInsideUnquotedValuesAsLiteral", html, 8 };
        }
        {
            string html = "<html><body><svg><foreignObject x=a/><![CDATA[x>" +
                string.Concat(Enumerable.Repeat("<div></div>", 64)) +
                "]]></foreignObject></svg></body></html>";
            yield return new object[] { "HtmlPreflightDoesNotTreatSlashInUnquotedAttributeAsSelfClosing", html, 16 };
        }
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
