using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRendering_UsesCanonicalHtmlListOrdinals() {
        const string html = "<ol start='9x'><li>First</li><li value='12junk'>Second</li><li>Third</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { "9. ", "First", "12. ", "Second", "13. ", "Third" },
            rendered.Text.Split('\n'));
    }
}
