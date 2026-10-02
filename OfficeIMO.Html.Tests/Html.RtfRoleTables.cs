using OfficeIMO.Html;
using OfficeIMO.Rtf;
using Xunit;
namespace OfficeIMO.Tests;
public class HtmlRtfRoleTables {
    [Fact]
    public void Rtf_ReportsUnsupportedAriaTableStructureAsApproximation() {
        const string html = "<div role='table'><p>Unstructured</p><div role='row'><div role='cell'>Value</div></div></div>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();

        Assert.Empty(result.RequireValue().Blocks.OfType<RtfTable>());
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void Rtf_ReportsActiveStylesheetsItDoesNotApply() {
        const string html = "<html><head><link rel='stylesheet' href='cid:layout.css'>"
            + "<link rel='stylesheet' disabled href='cid:disabled.css'>"
            + "<link rel='alternate stylesheet' title='alternate' href='cid:alternate.css'>"
            + "<link rel='stylesheet' type='text/less' href='cid:not-css.css'>"
            + "<style>p { color: red; }</style>"
            + "<style type='text/less'>p { color: blue; }</style>"
            + "</head><body><p>Content</p></body></html>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();

        Assert.Equal("cid:layout.css", Assert.Single(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "HtmlStylesheetLinkSkipped").Source);
        Assert.Equal(OfficeConversionLossKind.Omission, Assert.Single(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "HtmlStylesheetElementSkipped").LossKind);
    }

    [Fact]
    public void Rtf_AriaRoleTable_SyntheticNodesDoNotConsumeSourceLimits() {
        const string html = "<div role='table'><div role='row'><div role='cell'></div></div></div>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult(
            new HtmlToRtfOptions { MaxHtmlNodes = 6, MaxHtmlDepth = 5 });

        Assert.Single(result.RequireValue().Blocks.OfType<RtfTable>());
        HtmlRtfConversionLimitException exception = Assert.Throws<HtmlRtfConversionLimitException>(() =>
            HtmlConversionDocument.Parse(html).ToRtfDocumentResult(new HtmlToRtfOptions { MaxHtmlDepth = 4 }));
        Assert.Equal(nameof(HtmlToRtfOptions.MaxHtmlDepth), exception.LimitSource);
    }

    [Fact]
    public void Rtf_AriaZeroRowSpanStopsAtRowGroupBoundary() {
        const string html = "<div role='table'><div role='rowgroup'><div role='row'><span role='cell' aria-rowspan='0'>A</span><span role='cell'>B</span></div>"
            + "<div role='row'><span role='cell'>C</span></div></div><div role='rowgroup'><div role='row'><span role='cell'>D</span><span role='cell'>E</span></div></div></div>";
        var result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();
        var reopened = RtfDocument.Load(result.RequireValue().ToBytes());
        var table = Assert.Single(reopened.Blocks.OfType<RtfTable>());
        Assert.Equal(RtfTableCellMerge.First, table.Rows[0].Cells[0].VerticalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, table.Rows[1].Cells[0].VerticalMerge);
        Assert.NotEqual(RtfTableCellMerge.Continue, table.Rows[2].Cells[0].VerticalMerge);
    }
}
