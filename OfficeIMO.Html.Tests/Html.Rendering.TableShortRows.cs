using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlTables_ShortRowsMoveToTheNextPageWhole() {
        string rows = string.Concat(Enumerable.Range(0, 40).Select(index =>
            "<tr><td>Sku" + index.ToString("D2") + "<br>Line" + index.ToString("D2") + "</td><td>1</td></tr>"));
        string html = "<style>@page{size:300px 220px;margin:10px}table{width:100%;border-collapse:collapse}"
            + "th,td{font-size:10px;line-height:14px;padding:4px;border:1px solid #ccc}thead{display:table-header-group}</style>"
            + "<table><thead><tr><th>Item</th><th>Qty</th></tr></thead><tbody>" + rows + "</tbody></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.True(rendered.Pages.Count > 2);
        foreach (HtmlRenderPage page in rendered.Pages) {
            var texts = page.Visuals.OfType<HtmlRenderText>().Select(text => text.Text).ToList();
            foreach (string sku in texts.Where(text => text.StartsWith("Sku", StringComparison.Ordinal))) {
                Assert.Contains("Line" + sku.Substring(3), texts);
            }
        }
    }
}
