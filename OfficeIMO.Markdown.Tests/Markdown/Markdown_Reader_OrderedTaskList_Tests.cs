using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_OrderedTaskList_Tests {
    [Fact]
    public void Ordered_List_Items_Can_Be_Task_Items() {
        var md = """
1. [ ] Todo
2. [x] Done
""";
        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<ol", html, StringComparison.Ordinal);
        Assert.Contains("contains-task-list", html, StringComparison.Ordinal);
        Assert.Contains("task-list-item-checkbox", html, StringComparison.Ordinal);

        var round = doc.ToMarkdown().Trim();
        Assert.Contains("1. [ ] Todo", round, StringComparison.Ordinal);
        Assert.Contains("2. [x] Done", round, StringComparison.Ordinal);
    }

    [Fact]
    public void Ordered_Task_List_Markers_Require_Boundary_Whitespace_And_Support_Uppercase_X() {
        var md = "1. [X]\tUpper\n2. [ ]   Open\n3. [x]tight\n";

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var list = Assert.IsType<OrderedListBlock>(Assert.Single(doc.Blocks));

        Assert.Collection(
            list.Items,
            item => {
                Assert.True(item.IsTask);
                Assert.True(item.Checked);
                Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 2), item.MarkerSourceSpan);
                Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 6), item.TaskMarkerSourceSpan);
                Assert.Equal("Upper", item.Content.RenderMarkdown());
            },
            item => {
                Assert.True(item.IsTask);
                Assert.False(item.Checked);
                Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 2), item.MarkerSourceSpan);
                Assert.Equal(new MarkdownSourceSpan(2, 4, 2, 6), item.TaskMarkerSourceSpan);
                Assert.Equal("Open", item.Content.RenderMarkdown());
            },
            item => {
                Assert.False(item.IsTask);
                Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 2), item.MarkerSourceSpan);
                Assert.Null(item.TaskMarkerSourceSpan);
                Assert.Equal("[x]tight", InlinePlainText.Extract(item.Content));
            });
    }

    [Fact]
    public void ListExtras_Ordered_Task_Items_Preserve_Task_Marker_Source() {
        const string md = "a. [x] Alpha\n";
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.ListExtras = true;

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, options);
        var list = Assert.IsType<OrderedListBlock>(Assert.Single(doc.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Equal(MarkdownOrderedListMarkerStyle.LowerAlpha, list.MarkerStyle);
        Assert.True(item.IsTask);
        Assert.True(item.Checked);
        Assert.Equal("a.", item.MarkerText);
        Assert.Equal("[x]", item.TaskMarkerText);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 2), item.MarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 6), item.TaskMarkerSourceSpan);
        Assert.Equal("Alpha", item.Content.RenderMarkdown());
    }
}

