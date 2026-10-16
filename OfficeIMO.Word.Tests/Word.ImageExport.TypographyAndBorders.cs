using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordImageExportTests {
    [Fact]
    public void WordDocument_ImageTextUsesInheritedStyleAndDirectRunOverride() {
        using var document = WordDocument.Create();
        document.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.Append(new Style(
            new StyleName { Val = "Image heading" },
            new StyleRunProperties(new FontSize { Val = "41" }, new Bold(),
                new RunFonts { Ascii = "Arial" })) {
            Type = StyleValues.Paragraph, StyleId = "ImageHeading"
        });
        var paragraph = document.AddParagraph("Styled heading");
        paragraph.SetStyleId("ImageHeading");
        var direct = paragraph.AddText(" override");
        direct.Bold = false;
        direct._runProperties!.Bold = new Bold { Val = false };
        direct.FontSizePoints = 13.5;
        string svg = document.ToSvg();
        Assert.Contains("20.5", svg);
        var snapshot = document.CreateVisualSnapshot();
        var text = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>(),
            item => item.PlainText == "Styled heading override");
        Assert.Equal(20.5D, text.Runs[0].FontSize);
        Assert.True(text.Runs[0].Bold);
        Assert.Equal("Arial", text.Runs[0].FontFamily);
        Assert.Equal(13.5D, text.Runs[1].FontSize);
        Assert.False(text.Runs[1].Bold);
    }

    [Fact]
    public void WordDocument_ImageTableRetainsDirectTableBorderAndExplicitNone() {
        using var document = WordDocument.Create();
        var table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Bordered cell";
        table._tableProperties!.TableBorders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Color = "FF0000", Size = 16 },
            new BottomBorder { Val = BorderValues.Nil },
            new LeftBorder { Val = BorderValues.Nil },
            new RightBorder { Val = BorderValues.Nil });
        string svg = document.ToSvg();
        Assert.Contains("#FF0000", svg, StringComparison.OrdinalIgnoreCase);
    }
}
