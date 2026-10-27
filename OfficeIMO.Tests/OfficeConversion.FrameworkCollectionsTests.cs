using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeConversionFrameworkCollectionsTests {
    [Fact]
    public void HtmlTableRowGroupsPreserveCellPositionsAcrossRowspanZero() {
        const string html = "<table><thead><tr><th colspan='2'>Heading</th></tr></thead>"
            + "<tbody><tr><td rowspan='0'>Carry</td><td>First</td></tr><tr><td>Second</td></tr></tbody>"
            + "<tbody><tr><td>Other</td><td>Group</td></tr></tbody></table>";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(html));
        using var output = new MemoryStream();
        OfficeConverter.Convert(input, OfficeDocumentFormat.Html, output, OfficeDocumentFormat.Word);
        output.Position = 0;
        using var word = WordDocument.Load(output);
        var table = Assert.Single(word.Tables);
        Assert.Equal(4, table.Rows.Count);
        Assert.Equal("First", Text(1, 1));
        Assert.Equal("Second", Text(2, 1));
        Assert.Equal("Other", Text(3, 0));
        Assert.Equal("Group", Text(3, 1));
        string Text(int row, int column) => string.Join("", table.Rows[row].Cells[column].Paragraphs.Select(p => p.Text));
    }

    [Fact]
    public void SerializedShadowContentAndAssignedSlotsSurvivePdfRendering() {
        var html = HtmlConversionDocument.Parse("<div><span slot='content'>Assigned content</span>"
            + "<template shadowmode='open'><p>Shadow heading</p><slot name='content'></slot></template></div>");
        using var pdf = html.ToPdf(new HtmlToPdfOptions { ProjectSerializedShadowRoots = true });
        string text = pdf.ExtractText();
        Assert.Contains("Shadow heading", text, StringComparison.Ordinal);
        Assert.Contains("Assigned content", text, StringComparison.Ordinal);
    }
}
