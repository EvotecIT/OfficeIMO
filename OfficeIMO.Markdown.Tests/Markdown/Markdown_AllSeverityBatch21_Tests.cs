using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public class MarkdownAllSeverityBatch21Tests {
    [Theory]
    [InlineData(@"\\")]
    [InlineData(@"\#")]
    [InlineData(@"\+")]
    [InlineData(@"\-")]
    [InlineData(@"\.")]
    [InlineData(@"\!")]
    [InlineData(@"\(")]
    [InlineData(@"\}")]
    [InlineData(@"\~")]
    public void RawTableCellsPreserveLiteralBackslashes(string value) {
        TableBlock table = new();
        table.Headers.Add("Value");
        table.Rows.Add(new[] { value });

        string markdown = new MarkdownDoc().Add(table).ToMarkdown();
        MarkdownDoc reparsed = MarkdownReader.Parse(markdown);
        TableBlock reparsedTable = Assert.IsType<TableBlock>(Assert.Single(reparsed.Blocks));

        Assert.Equal(value, reparsedTable.Rows[0][0]);
    }
}
