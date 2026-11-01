using System.Text;
using System.Threading;
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownAllSeverityBatch12SecurityTests {
    [Fact]
    public void MarkdownReader_MixedEmphasisMarkersUseIndexedClosingRuns() {
        string markdown = "*outer "
            + string.Join(" ", Enumerable.Repeat("_inner", 4_000))
            + "*";

        MarkdownDocument document = MarkdownDocument.Parse(markdown);

        Assert.Single(document.Blocks);
        Assert.Contains("outer", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownReader_RejectsDetailsBeyondConfiguredNestingDepth() {
        string markdown = BuildNestedDetails(8);
        var options = new MarkdownLoadOptions { MaxNestingDepth = 4 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => MarkdownDocument.Parse(markdown, options));

        Assert.Contains("MaxNestingDepth (4)", exception.Message,
            StringComparison.Ordinal);
        Assert.Single(MarkdownDocument.Parse(
            BuildNestedDetails(2), options).Blocks);
    }

    [Theory]
    [InlineData(8, 4, false)]
    [InlineData(800, 64, false)]
    [InlineData(800, 64, true)]
    public void MarkdownReader_RejectsListsBeyondConfiguredNestingDepth(int levels, int maximumDepth, bool syntaxTree) {
        string markdown = BuildNestedList(levels);
        var options = new MarkdownLoadOptions { MaxNestingDepth = maximumDepth };

        InvalidDataException exception = RunOnSmallStack(() => Assert.Throws<InvalidDataException>(() => {
            if (syntaxTree) {
                MarkdownReader.ParseWithSyntaxTree(markdown, options);
            } else {
                MarkdownDocument.Parse(markdown, options);
            }
        }));

        Assert.Contains($"MaxNestingDepth ({maximumDepth})", exception.Message, StringComparison.Ordinal);
        Assert.Single(MarkdownDocument.Parse(BuildNestedList(3), options).Blocks);
    }

    [Fact]
    public void MarkdownReader_DeepListInQuoteKeepsItsNestingAndSourceLines() {
        const int levels = 48;
        string markdown = string.Concat(BuildNestedList(levels).Split('\n')
            .Where(line => line.Length > 0).Select(line => "> " + line + "\n"));

        MarkdownParseInfo parsed = MarkdownReader.ParseWithSyntaxTree(markdown);

        IMarkdownBlock block = Assert.Single(parsed.Document.Blocks);
        int depth = 1;
        var list = (UnorderedListBlock)Assert.IsType<QuoteBlock>(block).ChildBlocks[0];
        while (list.Items[0].NestedBlocks.Count > 0) {
            list = (UnorderedListBlock)list.Items[0].NestedBlocks[0];
            depth++;
        }
        Assert.Equal(levels, depth);
        Assert.Equal(levels, list.SourceSpan?.StartLine);
    }

    [Fact]
    public void MarkdownReader_DeeplyBracketedLinkKeepsOnlyTheInnermostLink() {
        const int depth = 5_000;
        string markdown = new string('[', depth) + "a" + string.Concat(Enumerable.Repeat("](u)", depth));

        ParagraphBlock paragraph = RunOnSmallStack(
            () => Assert.IsType<ParagraphBlock>(Assert.Single(MarkdownDocument.Parse(markdown).Blocks)));

        LinkInline link = Assert.IsType<LinkInline>(Assert.Single(paragraph.Inlines.Nodes, node => node is LinkInline));
        Assert.Equal("a", link.Text);
        Assert.Equal("u", link.Url);
    }

    [Fact]
    public void MarkdownReader_HtmlBlockWithManyTagStartsAvoidsSuffixCopies() {
        string markdown = "<div>" + new string('<', 50_000) + "</div>";

        MarkdownDocument document = MarkdownDocument.Parse(markdown);

        Assert.Single(document.Blocks);
    }

    [Fact]
    public void MarkdownToWord_RejectsDeepBlockAndInlineTrees() {
        var blockDocument = MarkdownDocument.Create();
        blockDocument.Add(CreateNestedDetailsBlock(8));
        var options = new WordMarkdownImportOptions { MaxNestingDepth = 4 };

        InvalidDataException blockException = Assert.Throws<InvalidDataException>(
            () => blockDocument.ToWord(options));
        Assert.Contains("block nesting", blockException.Message,
            StringComparison.Ordinal);

        var inlineDocument = MarkdownDocument.Create();
        inlineDocument.Add(new ParagraphBlock(CreateNestedInlineSequence(8)));
        InvalidDataException inlineException = Assert.Throws<InvalidDataException>(
            () => inlineDocument.ToWord(options));
        Assert.Contains("inline nesting", inlineException.Message,
            StringComparison.Ordinal);

        var safeDocument = MarkdownDocument.Create();
        safeDocument.Add(new ParagraphBlock(CreateNestedInlineSequence(2)));
        using var converted = safeDocument.ToWord(options);
        Assert.NotNull(converted);
    }

    [Fact]
    public void MarkdownRenderer_TableClipboardExportsNeutralizeSpreadsheetFormulas() {
        var options = new MarkdownRenderer.MarkdownRendererOptions {
            EnableTableCopyButtons = true
        };
        string shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml(
            "Security", options);

        Assert.Contains("function omdSpreadsheetSafeCell(value)", shell,
            StringComparison.Ordinal);
        Assert.Contains("replace(/[\\t\\r\\n]+/g, ' ')", shell,
            StringComparison.Ordinal);
        Assert.Contains("/^[=+\\-@]/.test(s)", shell,
            StringComparison.Ordinal);
        Assert.Equal(2, CountOccurrences(shell,
            "omdSpreadsheetSafeCell(omdCellText(c))"));
    }

    private static string BuildNestedList(int levels) {
        var markdown = new StringBuilder();
        for (int level = 0; level < levels; level++) {
            markdown.Append(' ', level * 2).Append("- item\n");
        }

        return markdown.ToString();
    }

    // A stack overflow ends the process, so parsing runs at a typical worker stack size.
    private static T RunOnSmallStack<T>(Func<T> action) {
        T? result = default;
        System.Runtime.ExceptionServices.ExceptionDispatchInfo? failure = null;
        var thread = new Thread(() => {
            try {
                result = action();
            } catch (Exception exception) {
                failure = System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception);
            }
        }, 1024 * 1024);
        thread.Start();
        thread.Join();
        failure?.Throw();
        return result!;
    }

    private static string BuildNestedDetails(int depth) {
        var markdown = new StringBuilder();
        for (int index = 0; index < depth; index++) {
            markdown.AppendLine("<details>");
            markdown.AppendLine("<summary>Level</summary>");
        }
        markdown.AppendLine("safe");
        for (int index = 0; index < depth; index++) {
            markdown.AppendLine("</details>");
        }
        return markdown.ToString();
    }

    private static DetailsBlock CreateNestedDetailsBlock(int depth) {
        IMarkdownBlock child = new ParagraphBlock(
            new InlineSequence().Text("safe"));
        for (int index = 0; index < depth; index++) {
            child = new DetailsBlock(new SummaryBlock("Level"),
                new[] { child });
        }
        return (DetailsBlock)child;
    }

    private static InlineSequence CreateNestedInlineSequence(int depth) {
        InlineSequence current = new InlineSequence().Text("safe");
        for (int index = 0; index < depth; index++) {
            var parent = new InlineSequence();
            parent.ReplaceItems(new IMarkdownInline[] {
                new BoldSequenceInline(current)
            });
            current = parent;
        }
        return current;
    }

    private static int CountOccurrences(string source, string value) {
        int count = 0;
        int offset = 0;
        while ((offset = source.IndexOf(value, offset,
                   StringComparison.Ordinal)) >= 0) {
            count++;
            offset += value.Length;
        }
        return count;
    }
}
