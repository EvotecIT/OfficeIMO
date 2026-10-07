using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void LegacyDoc_TableHierarchy_ConditionalInheritanceAndRegionPrecedence(bool derived, bool direct) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 1, WordTableStyle.TableGrid);
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style grid = styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid");
        grid.StyleParagraphProperties = new StyleParagraphProperties(new SpacingBetweenLines { After = "80" });
        grid.StyleRunProperties = new StyleRunProperties(new Color { Val = "00FF00" });
        grid.Append(new TableStyleProperties(new StyleParagraphProperties(new SpacingBetweenLines { After = "999" }),
            new StyleRunProperties(new Color { Val = "333333" })) { Type = TableStyleOverrideValues.WholeTable });
        grid.Append(new TableStyleProperties(new StyleParagraphProperties(new SpacingBetweenLines { After = "240" }),
            new StyleRunProperties(new Color { Val = "FF0000" })) { Type = TableStyleOverrideValues.FirstRow });
        if (derived) {
            styles.Append(new Style(new StyleName { Val = "Derived Conditional Grid" }, new BasedOn { Val = "TableGrid" },
                new TableStyleProperties(new StyleParagraphProperties(new SpacingBetweenLines { After = "360" }),
                    new StyleRunProperties(new Color { Val = "0000FF" })) { Type = TableStyleOverrideValues.FirstRow }) {
                Type = StyleValues.Table, StyleId = "DerivedConditionalGrid", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = "DerivedConditionalGrid" };
        }
        table._tableProperties!.TableLook = new TableLook { FirstRow = true, NoHorizontalBand = true, NoVerticalBand = true };
        WordParagraph header = table.Rows[0].Cells[0].Paragraphs[0];
        header.Text = "Header";
        if (direct) { header.ColorHex = "FF00FF"; header.LineSpacingAfter = 60; }
        table.Rows[1].Cells[0].Paragraphs[0].Text = "Body";
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        WordTable result = Assert.Single(restored.Tables);
        Assert.Equal(direct ? "FF00FF" : derived ? "0000FF" : "FF0000", result.Rows[0].Cells[0].Paragraphs[0].ColorHex);
        Assert.Equal(direct ? 60 : derived ? 360 : 240, result.Rows[0].Cells[0].Paragraphs[0].LineSpacingAfter);
        Assert.Equal("00FF00", result.Rows[1].Cells[0].Paragraphs[0].ColorHex);
        Assert.Equal(80, result.Rows[1].Cells[0].Paragraphs[0].LineSpacingAfter);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void LegacyDoc_TableHierarchy_ParagraphRunStylePrecedesTableRunDefaults(bool inheritedParagraph, bool direct) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid").StyleRunProperties =
            new StyleRunProperties(new Color { Val = "FF0000" });
        styles.Append(new Style(new StyleName { Val = "Cell Blue" }, new BasedOn { Val = "Normal" },
            new StyleRunProperties(new Color { Val = "0000FF" })) {
            Type = StyleValues.Paragraph, StyleId = "CellBlue", CustomStyle = true
        });
        if (inheritedParagraph) {
            styles.Append(new Style(new StyleName { Val = "Derived Cell Blue" }, new BasedOn { Val = "CellBlue" }) {
                Type = StyleValues.Paragraph, StyleId = "DerivedCellBlue", CustomStyle = true
            });
        }
        WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "Paragraph style";
        paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId {
            Val = inheritedParagraph ? "DerivedCellBlue" : "CellBlue"
        };
        if (direct) paragraph.ColorHex = "00FF00";
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        Assert.Equal(direct ? "00FF00" : "0000FF", restored.Tables[0].Rows[0].Cells[0].Paragraphs[0].ColorHex);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void LegacyDoc_TableHierarchy_MatchesWordIgnoringNormalTableChildProperties(int tableMode) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1, tableMode is 0 or 3 ? WordTableStyle.TableNormal : WordTableStyle.TableGrid);
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style normal = styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableNormal");
        normal.StyleRunProperties = new StyleRunProperties(new Color { Val = "FF0000" });
        normal.StyleParagraphProperties = new StyleParagraphProperties(new SpacingBetweenLines { Before = "180" });
        normal.GetFirstChild<StyleTableProperties>()!.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = "FFFF00" };
        normal.GetFirstChild<StyleTableProperties>()!.TableCellMarginDefault!.TopMargin =
            new TopMargin { Width = "144", Type = TableWidthUnitValues.Dxa };
        if (tableMode == 3) table._tableProperties!.TableStyle = null;
        if (tableMode == 2) {
            styles.Append(new Style(new StyleName { Val = "Normal Grid Child" }, new BasedOn { Val = "TableGrid" }) {
                Type = StyleValues.Table, StyleId = "NormalGridChild", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = "NormalGridChild" };
        }
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Inherited defaults";
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        WordTableCell cell = restored.Tables[0].Rows[0].Cells[0];
        Assert.Equal(string.Empty, cell.ShadingFillColorHex);
        Assert.Equal((short)0, cell.MarginTopWidth);
        Assert.Equal(string.Empty, cell.Paragraphs[0].ColorHex);
        Assert.Equal(0, cell.Paragraphs[0].LineSpacingBefore ?? 0);
    }
}
