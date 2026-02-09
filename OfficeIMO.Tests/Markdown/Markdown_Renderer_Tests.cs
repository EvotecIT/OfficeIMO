using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Renderer_Tests {
    [Fact]
    public void Reader_Disallows_Javascript_Links_ByDefault() {
        var md = "[x](javascript:alert(1))";
        var doc = MarkdownReader.Parse(md);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("javascript:", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal); // rendered as plain text
    }

    [Fact]
    public void Nested_Parsing_Respects_DisallowFileUrls() {
        var options = new MarkdownReaderOptions { DisallowFileUrls = true, HtmlBlocks = false, InlineHtml = false };
        string md = """
- outer
  > [x](file:///c:/test)
""";
        var doc = MarkdownReader.Parse(md, options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("file:", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Shell_Contains_UpdateContent_And_Mermaid_Bootstrap() {
        var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat");
        Assert.Contains("async function updateContent", shell, StringComparison.Ordinal);
        Assert.Contains("omdRoot", shell, StringComparison.Ordinal);
        Assert.Contains("mermaid.esm.min.mjs", shell, StringComparison.Ordinal);
        Assert.Contains("katex", shell, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("renderMathInElement", shell, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownRenderer_Adds_Mermaid_Hash_Attributes() {
        var md = "```mermaid\nflowchart LR\nA-->B\n```";
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md);
        Assert.Contains("class=\"mermaid\"", html, StringComparison.Ordinal);
        Assert.Contains("data-mermaid-hash", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Converts_Chart_Code_Fences_When_Enabled() {
        var md = "```chart\n{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}\n```";
        var opts = new MarkdownRendererOptions();
        opts.Chart.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);
        Assert.Contains("canvas", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("class=\"omd-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-chart-config-b64", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Shell_Contains_ChartJs_When_Enabled() {
        var opts = new MarkdownRendererOptions();
        opts.Chart.Enabled = true;

        var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat", opts);
        Assert.Contains("chart.umd", shell, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownRenderer_BaseHref_Is_Emitted_As_Base_Tag() {
        var opts = new MarkdownRendererOptions { BaseHref = "https://example.com/" };
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("[link](/x)", opts);
        Assert.Contains("<base href=\"https://example.com/\">", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Defaults_To_Hardened_External_Links() {
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("[x](https://example.com)");
        Assert.Contains("target=\"_blank\"", html, StringComparison.Ordinal);
        Assert.Contains("rel=\"", html, StringComparison.Ordinal);
        Assert.Contains("noopener", html, StringComparison.Ordinal);
        Assert.Contains("noreferrer", html, StringComparison.Ordinal);
        Assert.Contains("nofollow", html, StringComparison.Ordinal);
        Assert.Contains("ugc", html, StringComparison.Ordinal);
        Assert.Contains("referrerpolicy=\"no-referrer\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Can_Restrict_Absolute_Http_Links_To_Base_Origin() {
        var opts = new MarkdownRendererOptions { BaseHref = "https://example.com/" };
        opts.HtmlOptions.RestrictHttpLinksToBaseOrigin = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("[x](https://other.example.com/path)", opts);
        Assert.DoesNotContain("other.example.com", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Can_Restrict_Absolute_Http_Images_To_Base_Origin() {
        var opts = new MarkdownRendererOptions { BaseHref = "https://example.com/" };
        opts.HtmlOptions.RestrictHttpImagesToBaseOrigin = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("![alt](https://other.example.com/a.png)", opts);
        Assert.DoesNotContain("<img", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("omd-image-blocked", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Defaults_Emit_Image_Hardening_Attributes() {
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("![alt](https://example.com/a.png)");
        Assert.Contains("<img", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("loading=\"lazy\"", html, StringComparison.Ordinal);
        Assert.Contains("decoding=\"async\"", html, StringComparison.Ordinal);
        Assert.Contains("referrerpolicy=\"no-referrer\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Task_Lists_Emit_GithubLike_Classes() {
        var md = """
- [ ] Todo
- [x] Done
""";
        var doc = MarkdownReader.Parse(md);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("contains-task-list", html, StringComparison.Ordinal);
        Assert.Contains("task-list-item", html, StringComparison.Ordinal);
        Assert.Contains("task-list-item-checkbox", html, StringComparison.Ordinal);
    }
}
