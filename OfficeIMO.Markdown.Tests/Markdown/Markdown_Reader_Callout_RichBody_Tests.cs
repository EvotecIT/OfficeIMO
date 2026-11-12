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

        var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md, new MarkdownLoadOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtml(new MarkdownHtmlExportOptions { Style = MarkdownHtmlStyle.Plain, CssDelivery = MarkdownCssDelivery.None, BodyClass = null, Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment });

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

        var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md, new MarkdownLoadOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtml(new MarkdownHtmlExportOptions { Style = MarkdownHtmlStyle.Plain, CssDelivery = MarkdownCssDelivery.None, BodyClass = null, Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment });
        var markdown = doc.ToMarkdown();

        Assert.Contains("<strong>Use <strong>strong</strong> <a href=\"https://example.com\">links</a></strong>", html, StringComparison.Ordinal);
        Assert.Contains("> [!TIP] Use **strong** [links](https://example.com)", markdown, StringComparison.Ordinal);

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(doc.Blocks));
        Assert.Equal("Use strong links", callout.Title);
        Assert.Equal("Use **strong** [links](https://example.com)", callout.TitleInlines.RenderMarkdown());
    }

    [Fact]
    public void Callout_LazyContinuation_Stops_Before_Block_Starters() {
        const string md = """
> [!NOTE]
Body
# Next
""";

        var document = OfficeIMO.Markdown.MarkdownDocument.Parse(md);

        Assert.Collection(document.Blocks,
            block => {
                var callout = Assert.IsType<CalloutBlock>(block);
                var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(callout.ChildBlocks));
                Assert.Equal("Body", paragraph.Inlines.RenderMarkdown());
            },
            block => {
                var heading = Assert.IsType<HeadingBlock>(block);
                Assert.Equal("Next", heading.Inlines.RenderMarkdown());
            });
    }

    [Fact]
    public void Callout_Title_With_ImageOnly_Inline_Does_Not_Fall_Back_To_Kind_Label() {
        string md = """
> [!TIP] ![](/icon.svg)
> Body
""";

        var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md, new MarkdownLoadOptions { HtmlBlocks = false, InlineHtml = false });
        var html = doc.ToHtml(new MarkdownHtmlExportOptions { Style = MarkdownHtmlStyle.Plain, CssDelivery = MarkdownCssDelivery.None, BodyClass = null, Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment });

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

        var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md, new MarkdownLoadOptions { HtmlBlocks = false, InlineHtml = false });

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(doc.Blocks));
        Assert.Equal("Use linked code", callout.Title);
        Assert.Equal("Use [linked `code`](https://example.com)", callout.TitleInlines.RenderMarkdown());

        var html = doc.ToHtml(new MarkdownHtmlExportOptions { Style = MarkdownHtmlStyle.Plain, CssDelivery = MarkdownCssDelivery.None, BodyClass = null, Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment });
        Assert.Contains("<strong>Use <a href=\"https://example.com\">linked <code>code</code></a></strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Callout_Body_Is_Derived_From_ChildBlocks_When_BlockContent_Is_Available() {
        var paragraph = new ParagraphBlock(OfficeIMO.Markdown.MarkdownReader.ParseInlineText("fresh body"));
        var callout = new CalloutBlock("note", new InlineSequence().Text("Heads up"), new IMarkdownBlock[] { paragraph });

        Assert.Equal("Heads up", callout.Title);
        Assert.Equal("fresh body", callout.Body);
        Assert.Same(paragraph, Assert.Single(callout.ChildBlocks));
    }

    [Fact]
    public void Callout_Public_Structured_Constructor_Uses_ChildBlocks_As_Primary_Body() {
        var paragraph = new ParagraphBlock(OfficeIMO.Markdown.MarkdownReader.ParseInlineText("fresh body"));
        var list = new UnorderedListBlock();
        list.Items.Add(new ListItem(OfficeIMO.Markdown.MarkdownReader.ParseInlineText("first")));
        var callout = new CalloutBlock("note", "Heads up", new IMarkdownBlock[] {
            paragraph,
            list
        });

        Assert.Equal(2, callout.ChildBlocks.Count);
        Assert.Same(paragraph, callout.ChildBlocks[0]);
        Assert.Same(list, callout.ChildBlocks[1]);
        Assert.Equal("fresh body\n\n- first", callout.Body.Replace("\r\n", "\n"));
        Assert.Contains("> - first", ((IMarkdownBlock)callout).RenderMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void Callout_Public_Text_Constructor_Adapts_Body_To_ChildBlocks() {
        var callout = new CalloutBlock("note", "Heads up", "first line\nsecond line");

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(callout.ChildBlocks));
        Assert.Equal("first line  \nsecond line", paragraph.Inlines.RenderMarkdown().Replace("\r\n", "\n"));
        Assert.Equal("first line\nsecond line", callout.Body.Replace("\r\n", "\n"));
        Assert.Contains("first line<br/>second line", ((IMarkdownBlock)callout).RenderHtml(), StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownDoc_Callout_BodyBuilder_Adds_Structured_ChildBlocks() {
        var document = MarkdownDocument.Create()
            .Callout("warning", "Structured", body => body
                .P("Intro")
                .Ul(list => list.Item("first").Item("second"))
                .Code("text", "payload"));

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(document.Blocks));
        Assert.Collection(
            callout.ChildBlocks,
            block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.Equal(new[] { "first", "second" }, Assert.IsType<UnorderedListBlock>(block).Items.Select(item => item.Content.RenderMarkdown()).ToArray()),
            block => Assert.Equal("payload", Assert.IsType<CodeBlock>(block).Content));

        var markdown = document.ToMarkdown().Replace("\r\n", "\n");
        Assert.Contains("> [!WARNING] Structured", markdown, StringComparison.Ordinal);
        Assert.Contains("> - first", markdown, StringComparison.Ordinal);
        Assert.Contains("> ```text", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("System.", markdown, StringComparison.Ordinal);
    }

    public static IEnumerable<object[]> CalloutPlacementProfiles() {
        yield return new object[] { "officeimo" };
        yield return new object[] { "commonmark-with-callouts" };
    }

    [Theory]
    [MemberData(nameof(CalloutPlacementProfiles))]
    public void Callout_Header_Indented_Four_Columns_Stays_Indented_Code(string profile) {
        foreach (string markdown in new[] { "Text\n\n    > [!NOTE]\n    > example\n", "Text\n\n\t> [!NOTE]\n\t> example\n" }) {
            var document = MarkdownDocument.Parse(markdown, CreateCalloutPlacementOptions(profile));

            Assert.Collection(
                document.Blocks,
                block => Assert.IsType<ParagraphBlock>(block),
                block => Assert.Equal("> [!NOTE]\n> example", Assert.IsType<CodeBlock>(block).Content.Replace("\r\n", "\n")));
        }
    }

    [Theory]
    [MemberData(nameof(CalloutPlacementProfiles))]
    public void Callout_Inside_List_Item_Stays_In_The_Item(string profile) {
        var document = MarkdownDocument.Parse("- item\n\n  > [!NOTE]\n  > nested\n", CreateCalloutPlacementOptions(profile));

        var item = Assert.Single(Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks)).Items);
        var callout = Assert.IsType<CalloutBlock>(Assert.Single(item.NestedBlocks));
        Assert.Equal("note", callout.Kind);
        Assert.Equal("nested", Assert.IsType<ParagraphBlock>(Assert.Single(callout.ChildBlocks)).Inlines.RenderMarkdown());
    }

    [Theory]
    [MemberData(nameof(CalloutPlacementProfiles))]
    public void Adjacent_Callouts_Inside_List_Item_Are_Both_Kept(string profile) {
        var document = MarkdownDocument.Parse("- item\n\n  > [!NOTE]\n  > first\n  > [!TIP]\n  > second\n", CreateCalloutPlacementOptions(profile));

        var item = Assert.Single(Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks)).Items);
        Assert.Equal(new[] { "note", "tip" }, item.NestedBlocks.Select(block => Assert.IsType<CalloutBlock>(block).Kind).ToArray());
    }

    [Theory]
    [MemberData(nameof(CalloutPlacementProfiles))]
    public void Callout_Header_Inside_Quoted_Fence_In_List_Item_Stays_Code(string profile) {
        var document = MarkdownDocument.Parse("- item\n\n  > ```text\n  > [!NOTE]\n  > example\n  > ```\n", CreateCalloutPlacementOptions(profile));

        var item = Assert.Single(Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks)).Items);
        var quote = Assert.IsType<QuoteBlock>(Assert.Single(item.NestedBlocks));
        var code = Assert.IsType<CodeBlock>(Assert.Single(quote.ChildBlocks));
        Assert.Equal("[!NOTE]\nexample", code.Content.Replace("\r\n", "\n"));
    }

    [Theory]
    [MemberData(nameof(CalloutPlacementProfiles))]
    public void Callout_Header_Inside_Quoted_Fence_Stays_Code(string profile) {
        var document = MarkdownDocument.Parse("> ```text\n> [!NOTE]\n> example\n> ```\n", CreateCalloutPlacementOptions(profile));

        var quote = Assert.IsType<QuoteBlock>(Assert.Single(document.Blocks));
        var code = Assert.IsType<CodeBlock>(Assert.Single(quote.ChildBlocks));
        Assert.Equal("[!NOTE]\nexample", code.Content.Replace("\r\n", "\n"));
    }

    [Theory]
    [MemberData(nameof(CalloutPlacementProfiles))]
    public void Callout_After_Fence_Text_Starts_Its_Own_Block_When_Fenced_Code_Is_Disabled(string profile) {
        var options = CreateCalloutPlacementOptions(profile);
        options.FencedCode = false;

        var document = MarkdownDocument.Parse("> ```text\n> [!NOTE]\n> body\n", options);

        Assert.Collection(
            document.Blocks,
            block => Assert.IsType<QuoteBlock>(block),
            block => Assert.Equal("note", Assert.IsType<CalloutBlock>(block).Kind));
    }

    [Theory]
    [MemberData(nameof(CalloutPlacementProfiles))]
    public void Callout_After_Quoted_Html_Comment_With_Backticks_Starts_Its_Own_Block(string profile) {
        var options = CreateCalloutPlacementOptions(profile);
        options.HtmlBlocks = true;

        var document = MarkdownDocument.Parse("> <!--\n> ```\n> -->\n> [!NOTE]\n> body\n", options);

        Assert.Collection(
            document.Blocks,
            block => Assert.IsType<QuoteBlock>(block),
            block => Assert.Equal("note", Assert.IsType<CalloutBlock>(block).Kind));
    }

    [Theory]
    [MemberData(nameof(CalloutPlacementProfiles))]
    public void Callout_At_Top_Level_With_Up_To_Three_Spaces_Is_Still_A_Callout(string profile) {
        foreach (string markdown in new[] { "> [!NOTE]\n> body\n", "   > [!NOTE] Custom title\n   > body\n" }) {
            var document = MarkdownDocument.Parse(markdown, CreateCalloutPlacementOptions(profile));

            Assert.Equal("note", Assert.IsType<CalloutBlock>(Assert.Single(document.Blocks)).Kind);
        }
    }

    private static MarkdownLoadOptions CreateCalloutPlacementOptions(string profile) {
        if (profile == "officeimo") return new MarkdownLoadOptions();

        var options = MarkdownLoadOptions.CreateCommonMarkProfile();
        options.Tables = true;
        options.TaskLists = true;
        options.FrontMatter = true;
        options.Callouts = true;
        MarkdownReaderBuiltInExtensions.AddCallouts(options);
        return options;
    }
}

