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

        var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://evil.example", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Reference_Definitions_Inside_Indented_Code_Are_Ignored() {
        var md = """
    [ref]: https://evil.example/

[x][ref]
""";

        var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://evil.example", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Reference_Definitions_Outside_Code_Work() {
        var md = """
[ref]: https://example.com/

[x][ref]
""";

        var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("href=\"https://example.com/\"", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal);
    }
}
