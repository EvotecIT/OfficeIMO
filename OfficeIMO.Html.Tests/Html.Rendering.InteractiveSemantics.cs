using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_ClosedDialogsAndDetailsHideInactiveContent() {
        const string html = """
            <main>
              <p>Always visible</p>
              <dialog><p>Closed dialog content</p></dialog>
              <dialog open><p>Open dialog content</p></dialog>
              <details><summary>Closed summary</summary><p>Closed detail content</p></details>
              <details open><summary>Open summary</summary><p>Open detail content</p></details>
            </main>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        string text = string.Join(" ", rendered.Pages.SelectMany(page => page.Visuals)
            .OfType<HtmlRenderText>().Select(visual => visual.Text));

        Assert.Contains("Always visible", text);
        Assert.Contains("Open dialog content", text);
        Assert.Contains("Closed summary", text);
        Assert.Contains("Open summary", text);
        Assert.Contains("Open detail content", text);
        Assert.DoesNotContain("Closed dialog content", text);
        Assert.DoesNotContain("Closed detail content", text);
    }
}
