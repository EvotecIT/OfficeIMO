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

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(doc.Blocks));
        Assert.Collection(
            callout.ChildBlocks,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<UnorderedListBlock>(block),
            block => Assert.IsType<CodeBlock>(block));
        Assert.Equal("""
First line

- Item 1
- Item 2

```csharp
Console.WriteLine("x");
```
""".Replace("\r\n", "\n"), callout.Body.Replace("\r\n", "\n"));
    }

    [Fact]
    public void Callout_Title_Preserves_Inline_Markup() {
        string md = """
> [!TIP] Use **strong** [links](https://example.com)
> Body
""";

        var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var markdown = doc.ToMarkdown();

        Assert.Contains("<strong>Use <strong>strong</strong> <a href=\"https://example.com\">links</a></strong>", html, StringComparison.Ordinal);
        Assert.Contains("> [!TIP] Use **strong** [links](https://example.com)", markdown, StringComparison.Ordinal);

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(doc.Blocks));
        Assert.Equal("Use strong links", callout.Title);
        Assert.Equal("Use **strong** [links](https://example.com)", callout.TitleInlines.RenderMarkdown());
    }

    [Fact]
    public void Callout_Title_With_ImageOnly_Inline_Does_Not_Fall_Back_To_Kind_Label() {
        string md = """
> [!TIP] ![](/icon.svg)
> Body
""";

        var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<strong><img src=\"/icon.svg\" alt=\"\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<strong>Tip</strong>", html, StringComparison.Ordinal);

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(doc.Blocks));
        Assert.Equal(string.Empty, callout.Title);
        Assert.Equal("![](/icon.svg)", callout.TitleInlines.RenderMarkdown());
    }

    [Fact]
    public void Callout_Title_Plain_Text_Uses_Nested_Link_And_Code_Content() {
        string md = """
> [!NOTE] Use [linked `code`](https://example.com)
> Body
""";

        var doc = MarkdownReader.Parse(md, new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false });

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(doc.Blocks));
        Assert.Equal("Use linked code", callout.Title);
        Assert.Equal("Use [linked `code`](https://example.com)", callout.TitleInlines.RenderMarkdown());

        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<strong>Use <a href=\"https://example.com\">linked <code>code</code></a></strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Callout_Body_Is_Derived_From_ChildBlocks_When_BlockContent_Is_Available() {
        var paragraph = new ParagraphBlock(MarkdownReader.ParseInlineText("fresh body"));
        var callout = new CalloutBlock("note", new InlineSequence().Text("Heads up"), new IMarkdownBlock[] { paragraph });

        Assert.Equal("Heads up", callout.Title);
        Assert.Equal("fresh body", callout.Body);
        Assert.Same(paragraph, Assert.Single(callout.ChildBlocks));
    }
}

