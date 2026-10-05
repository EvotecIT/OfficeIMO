using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.Linq;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordBorderStyle.Dashed, false)]
    [InlineData(WordBorderStyle.Dashed, true)]
    [InlineData(WordBorderStyle.Dotted, false)]
    [InlineData(WordBorderStyle.Dotted, true)]
    [InlineData(WordBorderStyle.Double, false)]
    [InlineData(WordBorderStyle.Double, true)]
    public void SaveAsPdf_TableBorderPatterns_PreserveDirectAndInheritedStrokes(WordBorderStyle borderStyle, bool inherited) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2);
        if (inherited) {
            BorderValues style = borderStyle == WordBorderStyle.Dashed ? BorderValues.Dashed
                : borderStyle == WordBorderStyle.Dotted ? BorderValues.Dotted : BorderValues.Double;
            const string styleId = "PatternedTable";
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new Style(
                new StyleName { Val = "Patterned table" },
                new StyleTableProperties(new TableBorders(
                    new TopBorder { Val = style, Color = "FF0000", Size = 12U },
                    new LeftBorder { Val = style, Color = "FF0000", Size = 12U },
                    new BottomBorder { Val = style, Color = "FF0000", Size = 12U },
                    new RightBorder { Val = style, Color = "FF0000", Size = 12U },
                    new InsideHorizontalBorder { Val = style, Color = "FF0000", Size = 12U },
                    new InsideVerticalBorder { Val = style, Color = "FF0000", Size = 12U }))) {
                Type = StyleValues.Table, StyleId = styleId, CustomStyle = true
            });
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
        } else {
            table.StyleDetails!.SetBordersForAllSides(borderStyle, 12U, OfficeIMO.Drawing.OfficeColor.Red);
        }
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Patterned border";
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Patterned border", string.Concat(pdf.GetPages().Select(page => page.Text)));
        string raw = PdfOperatorSearchText.From(bytes);
        Assert.Contains("1 0 0 RG", raw, StringComparison.Ordinal);
        Assert.Contains("1.5 w", raw, StringComparison.Ordinal);
        if (borderStyle == WordBorderStyle.Double) {
            Assert.True(raw.Split(new[] { " S" }, StringSplitOptions.None).Length - 1 >= 12,
                "The 2-by-2 table must retain paired border strokes.");
        } else {
            Assert.Contains(borderStyle == WordBorderStyle.Dashed ? "[4.5 2.25] 0 d" : "[1.5 2.25] 0 d", raw,
                StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData(WordBorderStyle.Dashed, "[4.5 2.25] 0 d")]
    [InlineData(WordBorderStyle.Dotted, "[1.5 2.25] 0 d")]
    public void SaveAsPdf_CellBorderPatterns_PreserveSpecifiedStroke(WordBorderStyle borderStyle, string pattern) {
        using WordDocument document = WordDocument.Create();
        var cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Paragraphs[0].Text = "Cell border pattern";
        cell.Borders.TopStyle = borderStyle;
        cell.Borders.TopColorHex = "FF0000";
        cell.Borders.TopSize = 12U;
        string raw = PdfOperatorSearchText.From(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Contains(pattern, raw, StringComparison.Ordinal);
    }
}
