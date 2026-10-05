using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_TableSpacing_DocumentGridConditionalDefinitionIsInherited(bool derived) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 1, WordTableStyle.TableGrid);
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style grid = styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid");
        grid.Append(new TableStyleProperties(new TableStyleConditionalFormattingTableCellProperties(
            new Shading { Val = ShadingPatternValues.Clear, Fill = "FFFF00" })) {
            Type = TableStyleOverrideValues.FirstRow
        });
        if (derived) {
            styles.Append(new Style(new StyleName { Val = "Conditional Grid Child" }, new BasedOn { Val = "TableGrid" }) {
                Type = StyleValues.Table, StyleId = "ConditionalGridChild", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = "ConditionalGridChild" };
        }
        table._tableProperties!.TableLook = new TableLook { FirstRow = true, NoHorizontalBand = true, NoVerticalBand = true };
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Header";
        table.Rows[1].Cells[0].Paragraphs[0].Text = "Body";
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        WordTable result = Assert.Single(restored.Tables);
        Assert.Equal("FFFF00", result.Rows[0].Cells[0].ShadingFillColorHex);
        Assert.Equal(string.Empty, result.Rows[1].Cells[0].ShadingFillColorHex);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_TableSpacing_MissingGridDefinitionUsesLibraryFallback(bool derived) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid").Remove();
        if (derived) {
            styles.Append(new Style(new StyleName { Val = "Missing Grid Child" }, new BasedOn { Val = "TableGrid" }) {
                Type = StyleValues.Table, StyleId = "MissingGridChild", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = "MissingGridChild" };
        }
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Fallback";
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        WordTableCell cell = Assert.Single(restored.Tables).Rows[0].Cells[0];
        Assert.Equal(240, Assert.Single(cell.Paragraphs).LineSpacing);
        Assert.Equal(0, cell.Paragraphs[0].LineSpacingAfter);
        Assert.Equal(WordBorderStyle.Single, cell.Borders.TopStyle);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_TableSpacing_UsesDocumentTableGridDefinition(bool derived) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style grid = styles.Elements<Style>().Single(s => s.StyleId?.Value == "TableGrid");
        grid.StyleParagraphProperties = new StyleParagraphProperties(
            new SpacingBetweenLines { After = "180", Line = "360", LineRule = LineSpacingRuleValues.Auto });
        grid.StyleRunProperties = new StyleRunProperties(new Bold());
        grid.GetFirstChild<StyleTableProperties>()!.TableBorders!.TopBorder =
            new TopBorder { Val = BorderValues.Double, Size = 12, Color = "FF0000" };
        grid.GetFirstChild<StyleTableProperties>()!.Shading =
            new Shading { Val = ShadingPatternValues.Clear, Fill = "00FF00" };
        if (derived) {
            styles.Append(new Style(new StyleName { Val = "Modified Grid Child" }, new BasedOn { Val = "TableGrid" }) {
                Type = StyleValues.Table, StyleId = "ModifiedGridChild", CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = "ModifiedGridChild" };
        }
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Source grid";
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Doc)));
        WordTableCell cell = Assert.Single(restored.Tables).Rows[0].Cells[0];
        WordParagraph paragraph = Assert.Single(cell.Paragraphs);
        Assert.Equal(180, paragraph.LineSpacingAfter);
        Assert.Equal(360, paragraph.LineSpacing);
        Assert.True(paragraph.Bold);
        Assert.Equal(WordBorderStyle.Double, cell.Borders.TopStyle);
        Assert.Equal("FF0000", cell.Borders.TopColorHex);
        Assert.Equal("00FF00", cell.ShadingFillColorHex);
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    public void LegacyDoc_TableSpacing_TrailingPointBookmarkExcludesSyntheticSeparator(int controlDepth, bool authoredBlank) {
        using WordDocument document = WordDocument.Create();
        WordTable first = document.AddTable(1, 1);
        first.Rows[0].Cells[0].Paragraphs[0].Text = "First";
        first._table.Append(new BookmarkStart { Id = "71", Name = "TrailingPoint" });
        if (authoredBlank) document.AddParagraph();
        WordTable second = document.AddTable(1, 1);
        second.Rows[0].Cells[0].Paragraphs[0].Text = "Second";
        Body body = document._wordprocessingDocument!.MainDocumentPart!.Document.Body!;
        body.InsertBefore(new BookmarkStart { Id = "72", Name = "Second" }, second._table);
        body.InsertBefore(new BookmarkEnd { Id = "71" }, second._table);
        body.InsertAfter(new BookmarkEnd { Id = "72" }, second._table);
        for (int depth = 0; depth < controlDepth; depth++) {
            var content = new SdtContentBlock();
            foreach (var child in body.ChildElements.Where(e => e is not SectionProperties).ToArray()) {
                child.Remove();
                content.Append(child);
            }
            body.PrependChild(new SdtBlock(new SdtProperties(new SdtAlias { Val = "Boundary tables" }), content));
        }
        byte[] bytes = document.ToBytes(WordFileFormat.Doc);
        LegacyDocDocument model = LegacyDocDocument.Load(bytes);
        var point = Assert.Single(model.Bookmarks, b => b.Name == "TrailingPoint");
        int previousEnd = "First\a\a".Length;
        Assert.Equal(previousEnd, point.StartCharacter);
        Assert.Equal(previousEnd + (authoredBlank ? 1 : 0), point.EndCharacter);
        var next = Assert.Single(model.Bookmarks, b => b.Name == "Second");
        Assert.Equal(previousEnd + 1, next.StartCharacter);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes));
        Assert.Equal(2, reopened.Tables.Count);
    }
}
