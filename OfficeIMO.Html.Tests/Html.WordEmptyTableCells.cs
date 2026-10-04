using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordEmptyTableCellTests {
    [Theory]
    [InlineData("")]
    [InlineData("<!-- empty cell -->")]
    [InlineData("<table><tr><td>Nested value</td></tr></table>")]
    public void EmptyOrTableOnlyHtmlCellSavesAValidNativeCell(string content) {
        using WordDocument document = HtmlConversionDocument.Parse(
            "<p>Before table</p><table><tr><td>" + content + "</td><td>Neighbor</td></tr></table><p>After table</p>")
            .ToWordDocument();
        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);

        Assert.Empty(reopened.ValidateDocument());
        WordTable outer = reopened.Tables[0];
        Assert.NotEmpty(outer.Rows[0].Cells[0].Paragraphs);
        Assert.All(outer.Rows[0].Cells[0].Paragraphs, paragraph => Assert.Equal("", paragraph.Text));
        Assert.Equal(new[] { "Neighbor" }, outer.Rows[0].Cells[1].Paragraphs
            .Select(paragraph => paragraph.Text).Where(text => !string.IsNullOrEmpty(text)));
        Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "Before table");
        Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "After table");
    }
}
