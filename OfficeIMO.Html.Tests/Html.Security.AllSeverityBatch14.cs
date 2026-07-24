using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlAllSeverityBatch14SecurityTests {
    [Fact]
    public void SemanticProjectionDoesNotExposeIgnoredDescendantText() {
        HtmlSemanticDocument semantic = HtmlConversionDocument.Parse(
            "<p>visible<style>STYLE-SECRET</style><script>SCRIPT-SECRET</script>tail</p>"
        ).SemanticDocument;

        string text = string.Join(" ", semantic.Sections.SelectMany(section => section.Blocks).Select(block => block.Text));
        Assert.Contains("visible", text, StringComparison.Ordinal);
        Assert.Contains("tail", text, StringComparison.Ordinal);
        Assert.DoesNotContain("STYLE-SECRET", text, StringComparison.Ordinal);
        Assert.DoesNotContain("SCRIPT-SECRET", text, StringComparison.Ordinal);
    }
}
