using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, false, false)]
    [InlineData(0, true, false)]
    [InlineData(0, false, true)]
    [InlineData(0, true, true)]
    [InlineData(1, false, false)]
    [InlineData(1, true, false)]
    [InlineData(1, false, true)]
    [InlineData(1, true, true)]
    public void LegacyDoc_TableConditionalPrecedence_IntersectingRegionsFollowWordOrder(int regionPair, bool derived, bool direct) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2, WordTableStyle.TableGrid);
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style grid = styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid");
        TableStyleOverrideValues stronger = regionPair == 0 ? TableStyleOverrideValues.FirstRow : TableStyleOverrideValues.Band1Vertical;
        TableStyleOverrideValues weaker = regionPair == 0 ? TableStyleOverrideValues.FirstColumn : TableStyleOverrideValues.Band1Horizontal;
        var weakCondition = new TableStyleProperties(
            new StyleParagraphProperties(new SpacingBetweenLines { After = "360" }),
            new StyleRunProperties(new Color { Val = "0000FF" })) { Type = weaker };
        if (derived) {
            styles.Append(new Style(new StyleName { Val = "Intersecting Regions" }, new BasedOn { Val = "TableGrid" }, weakCondition) {
                Type = StyleValues.Table, StyleId = "IntersectingRegions", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = "IntersectingRegions" };
        } else {
            grid.Append(weakCondition);
        }
        grid.Append(new TableStyleProperties(
            new StyleParagraphProperties(new SpacingBetweenLines { After = "240" }),
            new StyleRunProperties(new Color { Val = "FF0000" })) { Type = stronger });
        table._tableProperties!.TableLook = new TableLook {
            FirstRow = regionPair == 0, FirstColumn = regionPair == 0,
            LastRow = false, LastColumn = false,
            NoHorizontalBand = regionPair == 0, NoVerticalBand = regionPair == 0
        };
        for (int row = 0; row < 2; row++) {
            for (int column = 0; column < 2; column++) table.Rows[row].Cells[column].Paragraphs[0].Text = $"R{row}C{column}";
        }
        if (direct) {
            table.Rows[0].Cells[0].Paragraphs[0].ColorHex = "FF00FF";
            table.Rows[0].Cells[0].Paragraphs[0].LineSpacingAfter = 60;
        }
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        WordTable result = Assert.Single(restored.Tables);
        Assert.Equal(direct ? "FF00FF" : "FF0000", result.Rows[0].Cells[0].Paragraphs[0].ColorHex);
        Assert.Equal(direct ? 60 : 240, result.Rows[0].Cells[0].Paragraphs[0].LineSpacingAfter);
        WordParagraph strongerOnly = result.Rows[regionPair == 0 ? 0 : 1].Cells[regionPair == 0 ? 1 : 0].Paragraphs[0];
        WordParagraph weakerOnly = result.Rows[regionPair == 0 ? 1 : 0].Cells[regionPair == 0 ? 0 : 1].Paragraphs[0];
        Assert.Equal("FF0000", strongerOnly.ColorHex);
        Assert.Equal(240, strongerOnly.LineSpacingAfter);
        Assert.Equal("0000FF", weakerOnly.ColorHex);
        Assert.Equal(360, weakerOnly.LineSpacingAfter);
    }
}
