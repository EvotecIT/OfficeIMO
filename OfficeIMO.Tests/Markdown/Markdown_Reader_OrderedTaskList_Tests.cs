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
        var doc = MarkdownReader.Parse(md);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<ol", html, StringComparison.Ordinal);
        Assert.Contains("contains-task-list", html, StringComparison.Ordinal);
        Assert.Contains("task-list-item-checkbox", html, StringComparison.Ordinal);

        var round = doc.ToMarkdown().Trim();
        Assert.Contains("1. [ ] Todo", round, StringComparison.Ordinal);
        Assert.Contains("2. [x] Done", round, StringComparison.Ordinal);
    }
}

