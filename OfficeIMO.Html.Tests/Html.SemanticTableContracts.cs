using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlSemanticTableContractTests {
    [Fact]
    public void RoleTableRetainsHeadersSpansAndSourceRows() {
        var source = HtmlConversionDocument.Parse("<div role='table' aria-label='Values'><div role='row'></div>"
            + "<div role='row'><span role='columnheader' aria-colspan='2'>Name</span></div>"
            + "<div role='rowgroup'><div role='row'><span role='cell'>A</span><span role='cell'>B</span></div></div></div>");
        var table = Assert.Single(source.SemanticDocument.Sections.SelectMany(s => s.Blocks), b => b.Kind == HtmlSemanticBlockKind.Table).Table!;
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(1, table.Rows[0].SourceRowIndex);
        Assert.Equal(2, table.Rows[1].SourceRowIndex);
        Assert.True(table.Rows[0].Cells[0].IsHeader);
        Assert.Equal(2, table.Rows[0].Cells[0].ColumnSpan);
        Assert.Equal(new[] { "A", "B" }, table.Rows[1].Cells.Select(c => c.Text));
    }

    [Fact]
    public void AuthoredCaptionRunsRemainSeparateFromFallbackTitle() {
        var authored = HtmlConversionDocument.Parse("<table><caption>Units <strong>SI</strong></caption><tr><td>m</td></tr></table>")
            .SemanticDocument.Sections.SelectMany(s => s.Blocks).Single(b => b.Table != null).Table!;
        Assert.Equal("Units SI", string.Concat(authored.CaptionRuns.Select(r => r.Text)));
        var fallback = HtmlConversionDocument.Parse("<h2>Units</h2><table><tr><td>m</td></tr></table>")
            .SemanticDocument.Sections.SelectMany(s => s.Blocks).Single(b => b.Table != null).Table!;
        Assert.Empty(fallback.CaptionRuns);
        Assert.Equal("Units", fallback.Caption);
    }

    [Fact]
    public void ImageSemanticHyperlinksUseThePreparedDocumentPolicy() {
        const string image = "<img src='data:image/png;base64,iVBORw0KGgo=' alt='Photo'>";
        var source = HtmlConversionDocument.Parse("<a href='https://example.test/photo'>" + image + "</a>");
        var resource = Assert.Single(source.SemanticDocument.Resources);
        Assert.Equal("https://example.test/photo", resource.Hyperlink);
        var unsafeSource = HtmlConversionDocument.Parse("<a href='javascript:alert(1)'>" + image + "</a>");
        Assert.Null(Assert.Single(unsafeSource.SemanticDocument.Resources).Hyperlink);
    }
}
