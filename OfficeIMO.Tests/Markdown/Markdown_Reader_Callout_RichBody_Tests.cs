using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Callout_RichBody_Tests {
    [Fact]
    public void Callout_Body_Is_Parsed_As_Markdown_Blocks() {
        string md = """
> [!NOTE] Title
> First line
>
> - Item 1
> - Item 2
>
> ```csharp
> Console.WriteLine("x");
> ```
""";

        var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("blockquote class=\"callout note\"", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<strong>Title</strong>", html, StringComparison.Ordinal);
        Assert.Contains("<ul", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<li>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Item 1", html, StringComparison.Ordinal);
        Assert.Contains("Item 2", html, StringComparison.Ordinal);
        Assert.Contains("language-csharp", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Console.WriteLine", html, StringComparison.Ordinal);
    }
}

