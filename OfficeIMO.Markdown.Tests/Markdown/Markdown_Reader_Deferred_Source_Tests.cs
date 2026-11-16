using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Reader_Deferred_Source_Tests {
    private const string RichMarkdown =
        "# Release notes\n" +
        "\n" +
        "Intro with **strong**, _emphasis_, `code`, a [link](https://example.com \"Title\") and ![logo](logo.png).\\\n" +
        "Escaped \\* star and &amp; entity.\n" +
        "\n" +
        "- [x] done item\n" +
        "  - nested item\n" +
        "\n" +
        "  continuation paragraph\n" +
        "1. ordered\n" +
        "\n" +
        "> quoted *text*\n" +
        "\n" +
        "> [!NOTE] Heads up\n" +
        "> callout body\n" +
        "\n" +
        "| Name | Value |\n" +
        "| :--- | ----: |\n" +
        "| a | `b` |\n" +
        "\n" +
        "```csharp\n" +
        "var x = 1;\n" +
        "```\n" +
        "\n" +
        "Term\n" +
        ": Definition\n" +
        "\n" +
        "Footnote ref[^1].\n" +
        "\n" +
        "[^1]: Footnote body.\n" +
        "\n" +
        "<div>raw html</div>\n" +
        "\n" +
        "***\n";

    public static IEnumerable<object?[]> Profiles() {
        yield return new object?[] { null };
        yield return new object?[] { MarkdownLoadOptions.MarkdownDialectProfile.CommonMark };
        yield return new object?[] { MarkdownLoadOptions.MarkdownDialectProfile.GitHubFlavoredMarkdown };
        yield return new object?[] { MarkdownLoadOptions.MarkdownDialectProfile.Portable };
    }

    [Theory]
    [MemberData(nameof(Profiles))]
    public void Parse_SourceInformation_Matches_SourceBacked_Parse(MarkdownLoadOptions.MarkdownDialectProfile? profile) {
        MarkdownLoadOptions? options = profile.HasValue ? MarkdownLoadOptions.CreateProfile(profile.Value) : null;
        var expected = MarkdownReader.ParseWithSyntaxTree(RichMarkdown, options);

        var document = MarkdownDocument.Parse(RichMarkdown, options);
        string html = document.ToHtml();
        string markdown = document.ToMarkdown();
        var objects = document.Descendants().Prepend(document).ToList();
        var expectedObjects = expected.Document.Descendants().Prepend(expected.Document).ToList();

        Assert.Equal(expectedObjects.Select(item => item.GetType()), objects.Select(item => item.GetType()));
        Assert.Equal(expectedObjects.Select(item => item.SourceSpan), objects.Select(item => item.SourceSpan));
        var parseInfo = Assert.IsType<MarkdownParseInfo>(document.ParseInfo);
        Assert.Same(document, parseInfo.Document);
        Assert.Equal(
            expected.FinalSyntaxTree.DescendantsAndSelf().Select(NodeSignature),
            parseInfo.FinalSyntaxTree.DescendantsAndSelf().Select(NodeSignature));
        for (int i = 0; i < objects.Count; i++) {
            Assert.Equal(
                NodeSignature(expected.FindFinalNodeForAssociatedObject(expectedObjects[i])),
                NodeSignature(parseInfo.FindFinalNodeForAssociatedObject(objects[i])));
        }

        Assert.Equal(expected.Document.ToHtml(), html);
        Assert.Equal(expected.Document.ToMarkdown(), markdown);
        Assert.Equal(html, document.ToHtml());
        Assert.Equal(markdown, document.ToMarkdown());
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(parseInfo);
    }

    [Fact]
    public void Parse_Typed_SourceSpans_Match_SourceBacked_Parse_When_Read_First() {
        var expected = MarkdownReader.ParseWithSyntaxTree(RichMarkdown).Document;

        var code = MarkdownDocument.Parse(RichMarkdown).DescendantObjectsOfType<CodeBlock>().Single();
        var item = MarkdownDocument.Parse(RichMarkdown).DescendantObjectsOfType<ListItem>().First();
        var link = MarkdownDocument.Parse(RichMarkdown).DescendantObjectsOfType<LinkInline>().Single();
        var heading = MarkdownDocument.Parse(RichMarkdown).DescendantObjectsOfType<HeadingBlock>().Single();

        Assert.Equal(expected.DescendantObjectsOfType<CodeBlock>().Single().ContentSourceSpan, code.ContentSourceSpan);
        Assert.Equal(expected.DescendantObjectsOfType<ListItem>().First().TaskMarkerSourceSpan, item.TaskMarkerSourceSpan);
        Assert.Equal(expected.DescendantObjectsOfType<LinkInline>().Single().UrlSourceSpan, link.UrlSourceSpan);
        Assert.Equal(expected.DescendantObjectsOfType<HeadingBlock>().Single().OpeningMarkerSourceSpan, heading.OpeningMarkerSourceSpan);
        Assert.NotNull(code.ContentSourceSpan);
        Assert.NotNull(link.UrlSourceSpan);
    }

    [Fact]
    public void Parse_SourceInformation_Describes_The_Parsed_Text_After_Edits() {
        var expectedList = MarkdownReader.ParseWithSyntaxTree(RichMarkdown).Document.Blocks.OfType<UnorderedListBlock>().First();
        var list = MarkdownDocument.Parse(RichMarkdown).Blocks.OfType<UnorderedListBlock>().First();
        var removed = list.Items[0];
        var added = new ListItem(new InlineSequence().Text("added"));

        list.Items.RemoveAt(0);
        list.Items.Add(added);
        var parseInfo = Assert.IsType<MarkdownParseInfo>(list.Document!.ParseInfo);

        Assert.Equal(expectedList.SourceSpan, list.SourceSpan);
        Assert.Equal(expectedList.Items[0].SourceSpan, removed.SourceSpan);
        Assert.Equal(expectedList.Items[0].MarkerSourceSpan, removed.MarkerSourceSpan);
        Assert.Null(added.SourceSpan);
        Assert.NotNull(parseInfo.FindFinalNodeForAssociatedObject(removed));
        Assert.Null(parseInfo.FindFinalNodeForAssociatedObject(added));
    }

    [Fact]
    public void Parse_Table_Cell_Inlines_Read_Through_The_Table_Match_SourceBacked_Parse() {
        const string markdown =
            "| Name | Notes |\n" +
            "| --- | --- |\n" +
            "| a | - [link](https://example.com) |\n";
        var expected = MarkdownReader.ParseWithSyntaxTree(markdown).Document.Blocks.OfType<TableBlock>().Single();

        var table = MarkdownDocument.Parse(markdown).Blocks.OfType<TableBlock>().Single();
        var link = table.RowInlines[0][1].Nodes.OfType<LinkInline>().Single();
        var expectedLink = expected.RowInlines[0][1].Nodes.OfType<LinkInline>().Single();

        Assert.NotNull(expectedLink.UrlSourceSpan);
        Assert.Equal(expectedLink.UrlSourceSpan, link.UrlSourceSpan);
        Assert.Equal(expectedLink.SourceSpan, link.SourceSpan);
    }

    [Fact]
    public void Parse_Table_Cell_Inlines_Keep_Parsed_Source_After_Edits() {
        const string markdown =
            "| Name | Notes |\n" +
            "| --- | --- |\n" +
            "| a | - [link](https://example.com) |\n";
        var expected = MarkdownReader.ParseWithSyntaxTree(markdown).Document.Blocks.OfType<TableBlock>().Single();
        var expectedLink = expected.RowInlines[0][1].Nodes.OfType<LinkInline>().Single();

        var cell = MarkdownDocument.Parse(markdown).Blocks.OfType<TableBlock>().Single().RowInlines[0][1];
        var parsedLink = cell.Nodes.OfType<LinkInline>().Single();
        var replacement = new LinkInline("other", "https://example.org", title: null);
        cell.ReplaceItems(cell.Nodes.Select(node => ReferenceEquals(node, parsedLink) ? replacement : node).ToList());

        Assert.Equal(expectedLink.UrlSourceSpan, parsedLink.UrlSourceSpan);
        Assert.NotNull(parsedLink.UrlSourceSpan);
        Assert.Null(replacement.UrlSourceSpan);
    }

    private static string NodeSignature(MarkdownSyntaxNode? node) =>
        node == null ? "<none>" : node.Kind + "|" + node.SourceSpan?.ToString() + "|" + node.Literal;
}
