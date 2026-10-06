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
    public void SaveAsPdf_TableBorderPatterns_PreserveDirectAndInheritedStrokes(WordBorderStyle borderStyle, bool inherited) {
        using WordDocument document = CreatePatternedTableDocument(borderStyle, inherited);
        WordTable table = document.Tables[0];
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Patterned border";
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Patterned border", string.Concat(pdf.GetPages().Select(page => page.Text)));
        string raw = PdfOperatorSearchText.From(bytes);
        Assert.Contains("1 0 0 RG", raw, StringComparison.Ordinal);
        Assert.Contains("1.5 w", raw, StringComparison.Ordinal);
        Assert.Contains(borderStyle == WordBorderStyle.Dashed ? "[4.5 2.25] 0 d" : "[1.5 2.25] 0 d", raw,
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_TableDoubleBorders_PreserveTwoTracksAndClearCellText(bool inherited) {
        using WordDocument document = CreatePatternedTableDocument(WordBorderStyle.Double, inherited);
        WordTable table = document.Tables[0];
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 2; column++) {
                var cell = table.Rows[row].Cells[column];
                cell.Paragraphs[0].Text = $"gyp{row}{column}";
                if (row == 1) cell.ShadingFillColorHex = "FFE6A0";
            }
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var page = pdf.GetPage(1);
        var lines = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue).Select(bounds => bounds!.Value).ToArray();
        double[] vertical = lines.Where(bounds => bounds.Width < .001D && bounds.Height > .001D)
            .Select(bounds => Math.Round(bounds.Left, 3)).Distinct().OrderBy(value => value).ToArray();
        double[] horizontal = lines.Where(bounds => bounds.Height < .001D && bounds.Width > .001D)
            .Select(bounds => Math.Round(bounds.Top, 3)).Distinct().OrderBy(value => value).ToArray();
        Assert.Equal(6, vertical.Length);
        Assert.Equal(6, horizontal.Length);
        Assert.Equal(3, vertical[3] - vertical[2], 3);
        var words = page.GetWords().Where(word => word.Text.StartsWith("gyp", StringComparison.Ordinal)).ToArray();
        Assert.Equal(4, words.Length);
        foreach (var word in words) {
            var box = word.BoundingBox;
            Assert.DoesNotContain(lines, line => line.Left - .75D < box.Right && line.Right + .75D > box.Left &&
                line.Bottom - .75D < box.Top && line.Top + .75D > box.Bottom);
        }
    }

    private static WordDocument CreatePatternedTableDocument(WordBorderStyle borderStyle, bool inherited) {
        WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2);
        if (inherited) {
            BorderValues style = borderStyle switch {
                WordBorderStyle.Dashed => BorderValues.Dashed,
                WordBorderStyle.Dotted => BorderValues.Dotted,
                WordBorderStyle.Double => BorderValues.Double,
                _ => throw new ArgumentOutOfRangeException(nameof(borderStyle))
            };
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
        return document;
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
