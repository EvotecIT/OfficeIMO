using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordTableLayoutMode.AutoFit, true, 72, 90D)]
    [InlineData(WordTableLayoutMode.AutoFit, false, 72, 176.4D)]
    [InlineData(WordTableLayoutMode.Fixed, true, 72, 72D)]
    [InlineData(WordTableLayoutMode.Fixed, false, 72, 72D)]
    [InlineData(WordTableLayoutMode.AutoFit, true, 216, 216D)]
    [InlineData(WordTableLayoutMode.AutoFit, false, 216, 216D)]
    [InlineData(WordTableLayoutMode.Fixed, true, 216, 216D)]
    [InlineData(WordTableLayoutMode.Fixed, false, 216, 216D)]
    public void SaveAsPdf_MixedPercentageRowsRetainTheOrderedContentAndAbsolutePreferences(
        WordTableLayoutMode layout, bool percentageFirst, int absolutePoints, double expectedWidth) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 3000, 3000);
        table.AddRow();
        table.LayoutMode = layout;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        for (int row = 0; row < 2; row++) {
            for (int column = 0; column < 2; column++) {
                WordTableCell cell = table.Rows[row].Cells[column];
                cell.Paragraphs[0].Text = column == 0 ? "SA SA SA SA SA SA SA SA" : "SB SB SB SB";
                ConfigurePartialColumnParagraph(cell.Paragraphs[0]);
                cell.MarginLeftWidth = 108; cell.MarginRightWidth = 108;
                cell.MarginTopWidth = 0; cell.MarginBottomWidth = 0;
                bool percentage = (row == column) == percentageFirst;
                cell._tableCellProperties!.TableCellWidth = new W.TableCellWidth {
                    Type = percentage ? W.TableWidthUnitValues.Pct : W.TableWidthUnitValues.Dxa,
                    Width = percentage ? "2500" : (absolutePoints * 20).ToString()
                };
            }
        }
        var fonts = new OfficeIMO.Pdf.PdfOptions();
        fonts.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600,
                Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = fonts
        }));
        var page = Assert.Single(pdf.GetPages());
        var letters = page.Letters.ToArray();
        string rowText = string.Concat(Enumerable.Repeat("SA", 8)) + string.Concat(Enumerable.Repeat("SB", 4));
        Assert.Equal(rowText + rowText, string.Concat(letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Select(letter => letter.Value)));
        double first = letters[0].StartBaseLine.X;
        double second = letters[Array.FindIndex(letters, letter => letter.Value == "B") - 1].StartBaseLine.X;
        Assert.InRange(second - first, expectedWidth - .05D, expectedWidth + .05D);
    }

    [Theory]
    [InlineData(2500, 1440, WordTableWidthUnit.Dxa, 41.9D)]
    [InlineData(3750, 1440, WordTableWidthUnit.Dxa, 62.85D)]
    [InlineData(2500, 0, WordTableWidthUnit.Auto, 80.9D)]
    [InlineData(3750, 0, WordTableWidthUnit.Auto, 121.35D)]
    [InlineData(2500, 0, WordTableWidthUnit.Dxa, 11.8D)]
    public void SaveAsPdf_FixedAutomaticPercentageFramesRetainUsableCellWidthsInNarrowContainers(
        int percentage, int otherWidth, WordTableWidthUnit otherType, double expectedFirstWidth) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 3000, 3000);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Paragraphs[0].Text = column == 0 ? "A" : "B";
            cell.WidthType = column == 0 ? WordTableWidthUnit.Pct : otherType;
            cell.Width = column == 0 ? percentage : otherWidth;
        }
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new OfficeIMO.Pdf.PageSize(164D, 792D),
            Margins = OfficeIMO.Pdf.PageMargins.Uniform(72D)
        }));
        var letters = Assert.Single(pdf.GetPages()).Letters;
        Assert.Equal("AB", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(expectedFirstWidth, letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, 2);
    }

    [Theory]
    [InlineData(false, 1250, 63.333D, 84D)]
    [InlineData(true, 1250, 75D, 100.364D)]
    [InlineData(true, 3750, 201D, 41.767D)]
    public void SaveAsPdf_FixedMixedPercentageColumnsRemoveSpacingFromCachedGridPreferences(
        bool explicitFrame, int percentage, double expectedFirstAdvance, double expectedSecondAdvance) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 1, 3, 120);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table._tableProperties!.TableWidth = new W.TableWidth {
            Type = explicitFrame ? W.TableWidthUnitValues.Dxa : W.TableWidthUnitValues.Auto,
            Width = explicitFrame ? "6000" : "0"
        };
        table.GridColumnWidth = new List<int> { 2000, 2000, 2000 };
        for (int column = 0; column < 3; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Paragraphs[0].Text = ((char)('A' + column)).ToString();
            cell.MarginLeftWidth = cell.MarginRightWidth = 54;
            cell.WidthType = column == 0 ? WordTableWidthUnit.Pct : column == 1 ? WordTableWidthUnit.Dxa : WordTableWidthUnit.Auto;
            cell.Width = column == 0 ? percentage : column == 1 ? 1440 : 0;
        }
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var letters = Assert.Single(pdf.GetPages()).Letters;
        Assert.Equal("ABC", string.Concat(letters.Select(letter => letter.Value)));
        Assert.InRange(letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, expectedFirstAdvance - .05D, expectedFirstAdvance + .05D);
        Assert.InRange(letters[2].StartBaseLine.X - letters[1].StartBaseLine.X, expectedSecondAdvance - .05D, expectedSecondAdvance + .05D);
    }

    [Fact]
    public void SaveAsPdf_ZeroSpacingPercentageAutoFitReservesFullBorderPaintInIntrinsicWidths() {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 1, 2, 0);
        table._tableProperties!.TableCellSpacing = new W.TableCellSpacing { Type = W.TableWidthUnitValues.Dxa, Width = "0" };
        table._tableProperties.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Paragraphs[0].Text = column == 0 ? "SA" : "SB";
            cell.MarginLeftWidth = cell.MarginRightWidth = 54;
            cell.WidthType = WordTableWidthUnit.Pct; cell.Width = 2500;
        }
        var fonts = new OfficeIMO.Pdf.PdfOptions();
        fonts.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 600,
                Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = fonts
        }));
        var letters = Assert.Single(pdf.GetPages()).Letters;
        Assert.Equal("SASB", string.Concat(letters.Select(letter => letter.Value)));
        Assert.InRange(letters[2].StartBaseLine.X - letters[0].StartBaseLine.X, 20.8D, 20.805D);
        Assert.Single(letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 3)).Distinct());
    }

    [Theory]
    [InlineData(WordTableWidthUnit.Auto, 3750, 125.35D)]
    [InlineData(WordTableWidthUnit.Nil, 3750, 125.35D)]
    [InlineData(WordTableWidthUnit.Auto, 5000, 100D)]
    [InlineData(WordTableWidthUnit.Nil, 5000, 100D)]
    public void SaveAsPdf_SingleColumnPercentageFrameRemovesBothOuterSpacingSharesFromCachedWidths(
        WordTableWidthUnit cachedType, int percentage, double expectedFrameWidth) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 2, 1, 120);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        table.GridColumnWidth = new List<int> { 2000 };
        for (int row = 0; row < 2; row++) {
            WordTableCell cell = table.Rows[row].Cells[0];
            cell.MarginLeftWidth = cell.MarginRightWidth = 54;
            cell.WidthType = row == 0 ? WordTableWidthUnit.Pct : cachedType;
            cell.Width = row == 0 ? percentage : 0;
            cell.Paragraphs[0].Text = "Alpha Bravo Charlie";
        }
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal("AlphaBravoCharlieAlphaBravoCharlie", string.Concat(page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).Select(letter => letter.Value)));
        var boxes = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(box => box.HasValue).Select(box => box!.Value).ToArray();
        Assert.InRange(boxes.Max(box => box.Right) - boxes.Min(box => box.Left),
            expectedFrameWidth - .05D, expectedFrameWidth + .05D);
    }
}
