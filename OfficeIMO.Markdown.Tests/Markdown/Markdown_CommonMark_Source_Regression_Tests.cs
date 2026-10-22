using System.Net;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_CommonMark_Source_Regression_Tests {
    private static readonly Lazy<Dictionary<int, Example>> Examples = new(LoadExamples);

    public static IEnumerable<object[]> CodeExamples() =>
        new[] { 6, 7, 112, 127, 131, 132, 133, 137, 139 }.Select(id => new object[] { id });

    public static IEnumerable<object[]> RawHtmlExamples() =>
        new[] { 21, 31, 148, 150, 156, 157, 158, 162, 163, 164, 165, 166, 167, 181, 187 }
            .Select(id => new object[] { id });

    [Theory]
    [MemberData(nameof(CodeExamples))]
    public void Official_Code_Examples_Preserve_Exact_Payload_And_Syntax(int id) {
        var example = Examples.Value[id];
        var result = MarkdownReader.ParseWithSyntaxTree(example.Markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var code = Assert.Single(Descendants(result.SyntaxTree)
            .Select(node => node.AssociatedObject).OfType<CodeBlock>().Distinct());
        var expected = CodePayload(example.Html);

        // The model excludes exactly one final line terminator; genuine blank lines remain.
        Assert.Equal(expected.Length == 0 ? expected : expected.Substring(0, expected.Length - 1), code.Content);
        Assert.Equal(expected, CodePayload(result.Document.ToHtmlFragment(PlainHtml())));
        Assert.Equal(expected, CodePayload(MarkdownReader.Parse(example.Markdown,
            MarkdownReaderOptions.CreateCommonMarkProfile()).ToHtmlFragment(PlainHtml())));
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Theory]
    [MemberData(nameof(RawHtmlExamples))]
    public void Official_RawHtml_Examples_Keep_Block_And_Preformatted_Line_Boundaries(int id) {
        var example = Examples.Value[id];
        var result = MarkdownReader.ParseWithSyntaxTree(example.Markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var actual = result.Document.ToHtmlFragment(PlainHtml());

        // These examples need only a final paragraph terminator / soft-break spelling allowance.
        // Interior raw HTML and preformatted newlines are compared literally.
        if (id == 187) {
            Assert.Equal(CommonMarkHtmlComparison.Normalize(example.Html), CommonMarkHtmlComparison.Normalize(actual));
        } else {
            Assert.Equal(example.Html.TrimEnd('\n'), actual.TrimEnd('\n'));
        }
        Assert.Equal(actual, MarkdownReader.Parse(example.Markdown,
            MarkdownReaderOptions.CreateCommonMarkProfile()).ToHtmlFragment(PlainHtml()));
        Assert.DoesNotContain(Descendants(result.SyntaxTree), node => node.AssociatedObject is HardBreakInline);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Theory]
    [InlineData(636)]
    [InlineData(637)]
    public void HardBreak_Continuation_Discards_Indent_And_Keeps_Original_Source_Column(int id) {
        var example = Examples.Value[id];
        var result = MarkdownReader.ParseWithSyntaxTree(example.Markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(result.Document.Blocks));
        Assert.Collection(paragraph.Inlines.Nodes,
            first => Assert.Equal("foo", Assert.IsType<MarkdownTextRun>(first).Text),
            marker => Assert.IsType<HardBreakInline>(marker),
            last => Assert.Equal("bar", Assert.IsType<MarkdownTextRun>(last).Text));
        var text = Assert.Single(Descendants(result.SyntaxTree), node => node.Literal == "bar");
        Assert.Equal(2, text.SourceSpan!.Value.StartLine);
        Assert.Equal(6, text.SourceSpan.Value.StartColumn);
        Assert.Equal("bar", example.Markdown.Substring(text.SourceSpan.Value.StartOffset!.Value, 3));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HardBreak_Before_Inline_Html_Remains_A_HardBreak(bool backslash) {
        var markdown = Examples.Value[187].Markdown.Replace("Foo\n", backslash ? "Foo\\\n" : "Foo  \n");
        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(result.Document.Blocks));
        Assert.Single(paragraph.Inlines.Nodes.OfType<HardBreakInline>());
        Assert.Empty(paragraph.Inlines.Nodes.OfType<SoftBreakInline>());
        var semantic = MarkdownReader.ParseSemantic(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        Assert.Equal(result.Document.ToHtmlFragment(PlainHtml()), semantic.ToHtmlFragment(PlainHtml()));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Theory]
    [InlineData(148, 0)]
    [InlineData(187, 0)]
    [InlineData(187, 1)]
    [InlineData(187, 2)]
    public void Semantic_And_SourceBacked_Html_SoftBreaks_Match_In_Containers(int id, int container) {
        var markdown = Examples.Value[id].Markdown;
        if (container == 1) {
            markdown = string.Join("\n", markdown.TrimEnd('\n').Split('\n').Select(line => "> " + line)) + "\n";
        } else if (container == 2) {
            markdown = "- " + markdown.TrimEnd('\n').Replace("\n", "\n  ") + "\n";
        }
        var source = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var semantic = MarkdownReader.ParseSemantic(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        Assert.Equal(source.Document.ToHtmlFragment(PlainHtml()), semantic.ToHtmlFragment(PlainHtml()));
        Assert.Empty(semantic.DescendantObjectsOfType<HardBreakInline>());
        Assert.NotEmpty(semantic.DescendantObjectsOfType<SoftBreakInline>());
        Assert.All(semantic.DescendantObjectsOfType<SoftBreakInline>(), node => Assert.Null(node.SourceSpan));
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(semantic);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void Escaped_Delimiter_Suffix_Does_Not_Change_Completed_Nested_Emphasis(int input) {
        var markdown = input == 3 ? "__bold _both___" : "**bold *both***";
        markdown += input switch {
            1 => " &amp; `[fail]literal[/]` \\*plain\\*",
            2 => " \\*plain\\*",
            3 => " &amp; `[fail]literal[/]` \\_plain\\_",
            4 => " \\\\\\*plain\\\\\\*",
            _ => string.Empty
        };
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.Tables = true;
        var source = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        foreach (var document in new[] { source.Document, MarkdownReader.Parse(markdown, options), MarkdownReader.ParseSemantic(markdown, options) }) {
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
            var bold = Assert.IsType<BoldSequenceInline>(paragraph.Inlines.Nodes[0]);
            Assert.Collection(bold.Inlines.Nodes,
                text => Assert.Equal("bold ", Assert.IsType<MarkdownTextRun>(text).Text),
                italic => Assert.Equal("both", Assert.IsType<MarkdownTextRun>(Assert.Single(Assert.IsType<ItalicSequenceInline>(italic).Inlines.Nodes)).Text));
            var html = document.ToHtmlFragment(PlainHtml());
            Assert.Contains("<strong>bold <em>both</em></strong>", html);
            if (input == 1) {
                Assert.Contains("</strong> &amp; <code>[fail]literal[/]</code> *plain*", html);
            }
        }
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(source);
    }

    [Fact]
    public void Even_Backslash_Run_Leaves_Emphasis_Markers_Active() {
        const string markdown = "\\\\**bold**";
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        foreach (var document in new[] { MarkdownReader.Parse(markdown, options), MarkdownReader.ParseWithSyntaxTree(markdown, options).Document, MarkdownReader.ParseSemantic(markdown, options) }) {
            Assert.Contains("\\<strong>bold</strong>", document.ToHtmlFragment(PlainHtml()));
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void List_Setext_Heading_Preserves_Soft_And_Explicit_Hard_Breaks(int breakKind) {
        var marker = breakKind == 1 ? "  " : breakKind == 2 ? "\\" : string.Empty;
        var markdown = "- Foo" + marker + "\n  <span>bar</span>  \n  ===\n";
        var source = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var semantic = MarkdownReader.ParseSemantic(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        foreach (var document in new[] { source.Document, semantic }) {
            var heading = Assert.Single(document.DescendantObjectsOfType<HeadingBlock>());
            Assert.Equal(breakKind == 0 ? 1 : 0, heading.Inlines.Nodes.OfType<SoftBreakInline>().Count());
            Assert.Equal(breakKind == 0 ? 0 : 1, heading.Inlines.Nodes.OfType<HardBreakInline>().Count());
            Assert.Equal("Foo", Assert.IsType<MarkdownTextRun>(heading.Inlines.Nodes[0]).Text);
        }
        Assert.Equal(source.Document.ToHtmlFragment(PlainHtml()), semantic.ToHtmlFragment(PlainHtml()));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(source);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Attributed_Indented_Paragraph_Trimming_Preserves_LineBreak_Tokens(int breakKind) {
        var marker = breakKind == 1 ? "  " : breakKind == 2 ? "\\" : string.Empty;
        var markdown = "{#lead .wide}\n    \tFoo" + marker + "\n    <span>bar</span>  \n";
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.GenericAttributes = true;
        var source = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var semantic = MarkdownReader.ParseSemantic(markdown, options);
        foreach (var document in new[] { source.Document, semantic }) {
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
            Assert.Equal("lead", paragraph.Attributes.ElementId);
            Assert.Contains("wide", paragraph.Attributes.Classes);
            Assert.Equal("Foo", Assert.IsType<MarkdownTextRun>(paragraph.Inlines.Nodes[0]).Text);
            Assert.Equal(breakKind == 0 ? 1 : 0, paragraph.Inlines.Nodes.OfType<SoftBreakInline>().Count());
            Assert.Equal(breakKind == 0 ? 0 : 1, paragraph.Inlines.Nodes.OfType<HardBreakInline>().Count());
        }
        Assert.Equal(source.Document.ToHtmlFragment(PlainHtml()), semantic.ToHtmlFragment(PlainHtml()));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(source);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Semantic_Quoted_Table_Uses_Table_Parser_Only_When_Enabled(bool tables) {
        const string markdown = "> | a | b |\n> | --- | --- |\n> | c | d |\n";
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.Tables = tables;
        var source = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var semantic = MarkdownReader.ParseSemantic(markdown, options);
        var quote = Assert.IsType<QuoteBlock>(Assert.Single(semantic.Blocks));
        if (tables) {
            Assert.IsType<TableBlock>(Assert.Single(quote.ChildBlocks));
            Assert.Contains("<td>c</td>", semantic.ToHtmlFragment(PlainHtml()));
        } else {
            Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
            Assert.DoesNotContain("<table>", semantic.ToHtmlFragment(PlainHtml()));
        }
        Assert.Equal(source.Document.ToHtmlFragment(PlainHtml()), semantic.ToHtmlFragment(PlainHtml()));
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(semantic);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Container_Code_Preserves_Tabs_After_Structural_Indentation(bool list) {
        var markdown = list ? "-     \tfoo\n" : ">     \tfoo\n";
        var source = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var semantic = MarkdownReader.ParseSemantic(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var code = Assert.Single(source.Document.DescendantObjectsOfType<CodeBlock>());
        Assert.Equal("\tfoo", code.Content);
        Assert.Equal("\tfoo\n", CodePayload(source.Document.ToHtmlFragment(PlainHtml())));
        Assert.Equal("\tfoo\n", CodePayload(semantic.ToHtmlFragment(PlainHtml())));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(source);
    }

    [Theory]
    [InlineData(0, 5, 2)]
    [InlineData(1, 5, 3)]
    [InlineData(2, 4, 3)]
    public void Quoted_Paragraph_Indentation_Preserves_Visual_Columns_And_Source_Offsets(int input, int column, int offset) {
        var markdown = input == 0 ? ">\tfoo\n" : input == 1 ? "> \tfoo\n" : ">  foo\n";
        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var text = Assert.Single(Descendants(result.SyntaxTree), node => node.AssociatedObject is MarkdownTextRun);
        Assert.Equal("foo", text.Literal);
        Assert.Equal(column, text.SourceSpan!.Value.StartColumn);
        Assert.Equal(column + 2, text.SourceSpan.Value.EndColumn);
        Assert.True(result.TryCreateSourceSlice(text, out var slice));
        Assert.Equal(offset, slice.StartOffset);
        Assert.Equal(offset + 2, slice.EndOffsetInclusive);
        Assert.Equal("foo", slice.Text);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void BuiltIn_Html_Controls_Keep_Their_Unicode_Labels() {
        var document = MarkdownReader.Parse("Note[^n].\n\n[^n]: body\n", MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var options = PlainHtml();
        options.ThemeToggle = true;
        options.GitHubFootnoteHtml = true;
        var html = document.ToHtmlFragment(options);
        Assert.Contains(char.ConvertFromUtf32(0x1F313), html);
        Assert.Contains("\u21a9", html);
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void Fenced_Code_Preserves_Real_Trailing_Blank_Lines(bool closed, int blankLines) {
        var content = "payload" + new string('\n', blankLines);
        var markdown = "```\n" + content + "\n" + (closed ? "```\n" : string.Empty);
        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var block = Assert.IsType<CodeBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal(content, block.Content);
        Assert.Equal(content + "\n", CodePayload(result.Document.ToHtmlFragment(PlainHtml())));
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void Fences_Distinguish_Empty_Content_From_Real_Blank_Lines(bool closed, int blankLines) {
        var payload = new string('\n', blankLines);
        var markdown = "```\n" + payload + (closed ? "```\n" : string.Empty);
        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        Assert.Equal(payload, CodePayload(result.Document.ToHtmlFragment(PlainHtml())));

        var written = result.Document.ToMarkdown();
        var reparsed = MarkdownReader.Parse(written, MarkdownReaderOptions.CreateCommonMarkProfile());
        Assert.Equal(payload, CodePayload(reparsed.ToHtmlFragment(PlainHtml())));
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unclosed_Fence_Without_Final_LineEnding_Still_Emits_One_Terminator(bool quoted) {
        var markdown = quoted ? "> ```\n> payload" : "```\npayload";
        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        Assert.Equal("payload\n", CodePayload(result.Document.ToHtmlFragment(PlainHtml())));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Quoted_Unclosed_Fence_Keeps_Explicit_Last_Blank_Source_Line(bool finalLineEnding) {
        var markdown = "> ```\n> payload\n>" + (finalLineEnding ? "\n" : string.Empty);
        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        Assert.Equal("payload\n\n", CodePayload(result.Document.ToHtmlFragment(PlainHtml())));
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
    }

    [Fact]
    public void Nested_List_Code_And_Following_Paragraph_Keep_Their_Boundaries() {
        const string markdown = "- item\n\n      chunk1\n        \n        chunk2\n\nOutside\n";
        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        Assert.IsType<UnorderedListBlock>(result.Document.Blocks[0]);
        Assert.IsType<ParagraphBlock>(result.Document.Blocks[1]);
        Assert.Equal("chunk1\n  \n  chunk2\n", CodePayload(result.Document.ToHtmlFragment(PlainHtml())));
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    private static string CodePayload(string html) {
        var match = Regex.Match(html, @"<code(?:\s[^>]*)?>([\s\S]*?)</code>");
        Assert.True(match.Success, "Expected a code element in the rendered fragment.");
        return WebUtility.HtmlDecode(match.Groups[1].Value);
    }

    private static IEnumerable<MarkdownSyntaxNode> Descendants(MarkdownSyntaxNode node) {
        yield return node;
        foreach (var child in node.Children) {
            foreach (var descendant in Descendants(child)) yield return descendant;
        }
    }

    private static HtmlOptions PlainHtml() => new() {
        Style = HtmlStyle.Plain,
        CssDelivery = CssDelivery.None,
        BodyClass = null,
        AutoHeadingIdentifiers = false,
        EscapeNonAsciiText = false,
        NormalizeUrlHostsToIdn = false
    };

    private static Dictionary<int, Example> LoadExamples() {
        for (var directory = new DirectoryInfo(Path.GetDirectoryName(SourceFile())!); directory != null; directory = directory.Parent) {
            var path = Path.Combine(directory.FullName, "OfficeIMO.Markdown.Tests", "Markdown", "Fixtures",
                "CommonMark", "commonmark-0.31.2-spec.json");
            if (!File.Exists(path)) continue;
            var examples = JsonSerializer.Deserialize<List<Example>>(File.ReadAllText(path),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            Assert.NotNull(examples);
            Assert.Equal(652, examples!.Count);
            return examples.ToDictionary(example => example.Id);
        }
        throw new DirectoryNotFoundException("Cannot locate the official CommonMark fixture from the test source directory.");
    }

    private static string SourceFile([System.Runtime.CompilerServices.CallerFilePath] string path = "") => path;

    private sealed class Example {
        public Example() { }
        [System.Text.Json.Serialization.JsonPropertyName("example")]
        public int Id { get; set; }
        public string Markdown { get; set; } = string.Empty;
        public string Html { get; set; } = string.Empty;
    }
}
