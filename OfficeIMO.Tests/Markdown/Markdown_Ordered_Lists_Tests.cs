using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownOrderedListsTests {
    [Fact]
    public void SeparateAdjacentOrderedListItems_InsertsBlankLineBetweenAdjacentOrderedItems() {
        const string markdown = """
            1. First check
            2. Second check
            """;

        var normalized = MarkdownOrderedLists.SeparateAdjacentOrderedListItems(markdown)
            .Replace("\r\n", "\n");

        Assert.Contains("1. First check\n\n2. Second check", normalized, StringComparison.Ordinal);
    }

    [Fact]
    public void SeparateAdjacentOrderedListItems_LeavesCleanMarkdownUnchanged() {
        const string markdown = """
            1. First check

            2. Second check
            """;

        var normalized = MarkdownOrderedLists.SeparateAdjacentOrderedListItems(markdown);

        Assert.Equal(markdown, normalized);
    }

    [Fact]
    public void SeparateAdjacentOrderedListItems_DoesNotSplitAdjacentOrderedItemsInsideFence() {
        const string markdown = """
            ```text
            1. First check
            2. Second check
            ```
            """;

        var normalized = MarkdownOrderedLists.SeparateAdjacentOrderedListItems(markdown)
            .Replace("\r\n", "\n");

        Assert.Contains("```text\n1. First check\n2. Second check\n```", normalized, StringComparison.Ordinal);
    }
}
