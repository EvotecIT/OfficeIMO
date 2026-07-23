using OfficeIMO.Markdown;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfMarkdownSecurityLimitTests {
    [Fact]
    public void MarkdownTableConversionRejectsConfiguredCellBudgetBeforeAllocation() {
        var table = new TableBlock();
        table.Headers.AddRange(new[] { "A", "B" });
        table.Rows.Add(new[] { "1", "2" });
        MarkdownDoc markdown = MarkdownDoc.Create().Add(table);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            markdown.ToRtfDocument(new MarkdownToRtfOptions { MaxTableCells = 3 }));

        Assert.Contains("exceeding the configured limit of 3", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownTableInsideListUsesTheSameCellBudget() {
        var table = new TableBlock();
        table.Headers.AddRange(new[] { "A", "B" });
        table.Rows.Add(new[] { "1", "2" });
        ListItem item = ListItem.Text("Item");
        item.NestedBlocks.Add(table);
        MarkdownDoc markdown = MarkdownDoc.Create().Add(new UnorderedListBlock { Items = { item } });

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            markdown.ToRtfDocument(new MarkdownToRtfOptions { MaxTableCells = 3 }));

        Assert.Contains("exceeding the configured limit of 3", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownListConversionCapsAuthoredRtfLevels() {
        string markdownText = "- zero\n  - one\n    - two\n      - three\n        - four\n";
        RtfConversionResult<OfficeIMO.Rtf.RtfDocument> result = MarkdownReader.Parse(markdownText).ToRtfDocumentResult(
            new MarkdownToRtfOptions { MaxListNestingDepth = 3 });

        Assert.All(result.Value.ListDefinitions, definition => Assert.InRange(definition.Levels.Count, 1, 3));
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "MDRTF018");
    }

    [Fact]
    public void MarkdownListConversionPreservesDepthFirstOrderWhenFlatteningCappedLists() {
        ListItem nested = ListItem.Text("AlphaChild");
        ListItem first = ListItem.Text("AlphaParent");
        first.NestedBlocks.Add(new UnorderedListBlock { Items = { nested } });
        ListItem second = ListItem.Text("BetaSibling");
        ListItem root = ListItem.Text("Root");
        root.NestedBlocks.Add(new UnorderedListBlock { Items = { first, second } });
        MarkdownDoc markdown = MarkdownDoc.Create().Add(new UnorderedListBlock { Items = { root } });

        string roundTrip = markdown.ToRtfDocument(new MarkdownToRtfOptions { MaxListNestingDepth = 1 }).ToMarkdown();

        int firstIndex = roundTrip.IndexOf("AlphaParent", StringComparison.Ordinal);
        int childIndex = roundTrip.IndexOf("AlphaChild", StringComparison.Ordinal);
        int secondIndex = roundTrip.IndexOf("BetaSibling", StringComparison.Ordinal);
        Assert.True(firstIndex >= 0 && childIndex > firstIndex && secondIndex > childIndex, roundTrip);
    }

    [Fact]
    public void MarkdownListConversionPreservesTaskMarkersWhenFlatteningCappedLists() {
        ListItem root = ListItem.Text("Root");
        root.NestedBlocks.Add(new UnorderedListBlock {
            Items = {
                ListItem.Task("Finished", done: true),
                ListItem.Task("Pending")
            }
        });
        MarkdownDoc markdown = MarkdownDoc.Create().Add(new UnorderedListBlock { Items = { root } });

        RtfDocument rtf = markdown.ToRtfDocument(new MarkdownToRtfOptions { MaxListNestingDepth = 1 });
        string plainText = string.Join("\n", rtf.Paragraphs.Select(paragraph => paragraph.ToPlainText()));
        string roundTrip = rtf.ToMarkdown();

        Assert.Contains("[x] Finished", plainText, StringComparison.Ordinal);
        Assert.Contains("[ ] Pending", plainText, StringComparison.Ordinal);
        Assert.Contains("\\[x\\] Finished", roundTrip, StringComparison.Ordinal);
        Assert.Contains("\\[ \\] Pending", roundTrip, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownListConversionPreservesOrderedMarkersWhenFlatteningCappedLists() {
        ListItem root = ListItem.Text("Root");
        root.NestedBlocks.Add(new OrderedListBlock {
            Start = 3,
            Items = {
                ListItem.Text("Third"),
                ListItem.Text("Fourth")
            }
        });
        MarkdownDoc markdown = MarkdownDoc.Create().Add(new UnorderedListBlock { Items = { root } });

        RtfDocument rtf = markdown.ToRtfDocument(new MarkdownToRtfOptions { MaxListNestingDepth = 1 });
        string plainText = string.Join("\n", rtf.Paragraphs.Select(paragraph => paragraph.ToPlainText()));

        Assert.Contains("3. Third", plainText, StringComparison.Ordinal);
        Assert.Contains("4. Fourth", plainText, StringComparison.Ordinal);
    }

    [Fact]
    public void CodeBlockInfoBookmarkPayloadIsBounded() {
        string infoString = new string('x', 2_000);
        MarkdownDoc markdown = MarkdownDoc.Create().Add(new CodeBlock(infoString, "value"));

        string roundTrip = markdown.ToRtfDocument().ToMarkdown();

        Assert.Contains(new string('x', 1_024), roundTrip, StringComparison.Ordinal);
        Assert.DoesNotContain(new string('x', 1_025), roundTrip, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    public void MarkdownConversionRejectsNonPositiveResourceLimits(int maxDepth, int maxCells) {
        var options = new MarkdownToRtfOptions {
            MaxListNestingDepth = maxDepth,
            MaxTableCells = maxCells
        };

        Assert.Throws<ArgumentOutOfRangeException>(() => MarkdownDoc.Create().ToRtfDocument(options));
    }
}
