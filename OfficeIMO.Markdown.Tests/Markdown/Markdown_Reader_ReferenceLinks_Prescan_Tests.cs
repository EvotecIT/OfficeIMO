using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_ReferenceLinks_Prescan_Tests {
    [Fact]
    public void Reference_Definitions_Inside_Fenced_Code_Are_Ignored() {
        var md = """
```bash
[ref]: https://evil.example/
```

[x][ref]
""";

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://evil.example", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Reference_Definitions_Inside_Indented_Code_Are_Ignored() {
        var md = """
    [ref]: https://evil.example/

[x][ref]
""";

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://evil.example", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Quoted_Reference_Definitions_Inside_Indented_Code_Are_Ignored() {
        var md = """
    > [ref]: https://evil.example/

[x][ref]
""";

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://evil.example", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("[x][ref]", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reference_Definitions_Outside_Code_Work() {
        var md = """
[ref]: https://example.com/

[x][ref]
""";

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("href=\"https://example.com/\"", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reference_PreScan_Uses_ListExtras_For_Paragraph_Interrupts() {
        var md = """
[x]
a. item
[x]: /url
""";
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.ListExtras = true;

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("href=\"/url\"", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<ol type=\"a\">", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reference_PreScan_Does_Not_Register_Quoted_Definition_Inside_Open_Quoted_Paragraph() {
        var md = """
[x]

> intro
> [x]: /url
""";
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"/url\"", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("[x]", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reference_PreScan_Treats_Custom_Containers_As_Paragraph_Breaks() {
        var md = """
[x]
::: note
body
:::
[x]: /url
""";
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.CustomContainers = true;

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("href=\"/url\"", html, StringComparison.OrdinalIgnoreCase);
        Assert.IsType<CustomContainerBlock>(doc.Blocks[1]);
    }
}
