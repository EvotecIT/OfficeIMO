using System.Threading;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    // A stack overflow kills the test host, so these inputs stay inside the depth limits
    // and run on the smallest default stack a supported host gives a worker thread.
    [Theory]
    [InlineData("<div>", "</div>", 200)]
    [InlineData("<div style='display:flex;flex-direction:column'>", "</div>", 200)]
    [InlineData("<span>", "</span>", 200)]
    [InlineData("<div style='display:contents'>", "</div>", 200)]
    [InlineData("<ul><li>", "</li></ul>", 100)]
    public void HtmlRender_NestingWithinTheDepthLimitsCompletesOnASmallCallerStack(string open, string close, int levels) {
        string html = string.Concat(Enumerable.Repeat(open, levels)) + "<span>DeepLeaf</span>" + string.Concat(Enumerable.Repeat(close, levels));
        HtmlRenderDocument? rendered = null;
        Exception? failure = null;
        var thread = new Thread(() => {
            try { rendered = HtmlRenderTestDriver.Render(html); } catch (Exception exception) { failure = exception; }
        }, 512 * 1024);
        thread.Start();
        thread.Join();

        Assert.Null(failure);
        Assert.Contains("DeepLeaf", rendered!.Text, StringComparison.Ordinal);
    }
}
