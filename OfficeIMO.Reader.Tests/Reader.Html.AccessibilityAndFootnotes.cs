using OfficeIMO.Reader;
using OfficeIMO.Reader.Html;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderHtmlAccessibilityAndFootnotesTests {
    [Fact]
    public void DocumentReaderHtml_LinkNamesPreferAriaThenVisibleTextBeforeTitle() {
        const string html = """
<p><a href="report" aria-label="Download annual report">Download</a></p>
<p><a href="guide" title="Open guide">Read more</a></p>
<span id="label" aria-label="Download labelled report">ignored</span>
<p><a href="labelled" aria-labelledby="label"></a></p>
""";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "links.html");

        Assert.Equal(new[] { "Download annual report", "Read more", "Download labelled report" }, result.Links.Select(link => link.Text));
    }

    [Fact]
    public void DocumentReaderHtml_ProjectsSourceLessAriaImageAsVisualWithoutAsset() {
        const string html = "<div role=\"img\" aria-label=\"Sales chart\"></div>";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "aria-image.html");
        ReaderVisual visual = Assert.Single(result.Visuals);

        Assert.Empty(result.Assets);
        Assert.Equal("image", visual.Kind);
        Assert.Equal("div", visual.Language);
        Assert.Equal("Sales chart", visual.Content);
        Assert.Null(visual.SourceName);
        Assert.Equal("image", visual.Location!.SourceBlockKind);
    }

    [Fact]
    public void DocumentReaderHtml_ProjectsGroupedEpubFootnotesAsFootnoteBlocks() {
        const string html = """
<p>Text<a epub:type="noteref" role="doc-noteref" href="#fn:1">1</a>.</p>
<ol epub:type="footnotes">
  <li id="fn:1">
    <p>Grouped <strong>footnote</strong>.</p>
    <a epub:type="backlink" role="doc-backlink" href="#ref:1">back</a>
  </li>
</ol>
""";

        OfficeDocumentReadResult result = HtmlReaderAdapter.ReadContentDocument(html, "grouped-footnotes.xhtml");
        OfficeDocumentBlock footnote = Assert.Single(result.Blocks, block => block.Kind == "footnote");

        Assert.Equal("Grouped footnote.", footnote.Text);
        Assert.DoesNotContain(result.Blocks, block => block.Kind == "list-item" && block.Text.Contains("Grouped", StringComparison.Ordinal));
        Assert.Contains("officeimo.html.footnotes", result.CapabilitiesUsed);
    }
}
