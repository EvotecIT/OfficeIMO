using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, 1, false)]
    [InlineData(false, 1, true)]
    [InlineData(false, 2, false)]
    [InlineData(false, 2, true)]
    [InlineData(true, 1, false)]
    [InlineData(true, 1, true)]
    [InlineData(true, 2, false)]
    [InlineData(true, 2, true)]
    public void SaveAsPdf_TableBandGroupsKeepFillAndTypographyTogether(bool vertical, int size, bool edge) {
        using WordDocument source = WordDocument.Create();
        WordTable table = source.AddTable(6, 6, WordTableStyle.TableGrid);
        var definition = new W.Style(new W.StyleName { Val = "Grouped bands" }, new W.StyleTableProperties(
            new W.TableStyleRowBandSize { Val = size }, new W.TableStyleColumnBandSize { Val = size })) {
            StyleId = "GroupedBands", Type = W.StyleValues.Table
        };
        for (int band = 1; band <= 2; band++) {
            definition.AppendChild(new W.TableStyleProperties(
                new W.RunPropertiesBaseStyle(new W.Color { Val = band == 1 ? "800000" : "000080" },
                    new W.Bold { Val = band == 2 }),
                new W.TableStyleConditionalFormattingTableCellProperties(new W.Shading {
                    Val = W.ShadingPatternValues.Clear, Fill = band == 1 ? "FFE0E0" : "E0E0FF"
                })) { Type = vertical ? (band == 1 ? W.TableStyleOverrideValues.Band1Vertical : W.TableStyleOverrideValues.Band2Vertical)
                    : (band == 1 ? W.TableStyleOverrideValues.Band1Horizontal : W.TableStyleOverrideValues.Band2Horizontal) });
        }
        source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.AppendChild(definition);
        table._tableProperties!.TableStyle = new W.TableStyle { Val = "GroupedBands" };
        table.ConditionalFormattingFirstRow = !vertical && edge;
        table.ConditionalFormattingFirstColumn = vertical && edge;
        table.ConditionalFormattingNoHorizontalBand = vertical;
        table.ConditionalFormattingNoVerticalBand = !vertical;
        for (int r = 0; r < 6; r++) for (int c = 0; c < 6; c++) {
            table.Rows[r].Cells[c].Paragraphs[0].Text = $"R{r}C{c}";
        }
        // Direct formatting still wins over either band.
        table.Rows[4].Cells[4].ShadingFillColorHex = "00FF00";
        table.Rows[4].Cells[4].Paragraphs[0].ColorHex = "008000";
        table.Rows[4].Cells[4].Paragraphs[0].Bold = false;
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        using var model = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false });
        using var pdf = PdfPigDocument.Open(model.ToBytes());
        var page = pdf.GetPage(1);
        for (int r = 0; r < 6; r++) for (int c = 0; c < 6; c++) {
            int index = vertical ? c : r;
            if (edge && index == 0) continue;
            bool first = (index - (edge ? 1 : 0)) / size % 2 == 0;
            bool direct = r == 4 && c == 4;
            var word = Assert.Single(page.GetWords(), item => item.Text == $"R{r}C{c}");
            var letter = word.Letters[0];
            var expectedFill = direct ? (0D, 1D, 0D) : first ? (1D, 224D / 255D, 224D / 255D) : (224D / 255D, 224D / 255D, 1D);
            var expectedText = direct ? (0D, 128D / 255D, 0D) : first ? (128D / 255D, 0D, 0D) : (0D, 0D, 128D / 255D);
            var actualText = letter.Color.ToRGBValues();
            Assert.Equal(expectedText.Item1, actualText.r, 3);
            Assert.Equal(expectedText.Item2, actualText.g, 3);
            Assert.Equal(expectedText.Item3, actualText.b, 3);
            Assert.Equal(!first && !direct, letter.FontName.Contains("Bold", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(page.Paths, path => path.IsFilled && MatchesBandColor(path.FillColor.ToRGBValues(), expectedFill) &&
                path.GetBoundingRectangle() is { } bounds && bounds.Left <= letter.StartBaseLine.X &&
                bounds.Right >= letter.StartBaseLine.X && bounds.Bottom <= letter.StartBaseLine.Y && bounds.Top >= letter.StartBaseLine.Y);
        }
        document.Tables[0].ConditionalFormattingNoHorizontalBand = true;
        document.Tables[0].ConditionalFormattingNoVerticalBand = true;
        using var unbandedModel = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false });
        using var unbanded = PdfPigDocument.Open(unbandedModel.ToBytes());
        var unbandedPage = unbanded.GetPage(1);
        Assert.Equal(36, unbandedPage.GetWords().Count());
        Assert.DoesNotContain(unbandedPage.Paths, path => path.IsFilled &&
            (MatchesBandColor(path.FillColor.ToRGBValues(), (1D, 224D / 255D, 224D / 255D)) ||
             MatchesBandColor(path.FillColor.ToRGBValues(), (224D / 255D, 224D / 255D, 1D))));
    }

    private static bool MatchesBandColor((double r, double g, double b) actual, (double r, double g, double b) expected) =>
        Math.Abs(actual.r - expected.r) < .00051D && Math.Abs(actual.g - expected.g) < .00051D && Math.Abs(actual.b - expected.b) < .00051D;
}
