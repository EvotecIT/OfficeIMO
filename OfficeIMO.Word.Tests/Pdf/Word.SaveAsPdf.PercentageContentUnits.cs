using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(600, false, 75.603D)]
    [InlineData(603, false, 75.903D)]
    [InlineData(600, true, 140.4D)]
    [InlineData(603, true, 141D)]
    public void SaveAsPdf_AutomaticPercentageTextCellsUseSourceWidthUnitsBeforeWrappingTolerance(
        ushort glyphAdvance, bool phrase, double expectedFirstColumnWidth) {
        using WordDocument source = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(source, 1440, 1440);
        table.LayoutMode = WordTableLayoutMode.AutoFit;
        table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Auto, Width = "0" };
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Paragraphs[0].Text = column == 0 ? "SA" : phrase ? "SB SB" : "SB";
            ConfigurePartialColumnParagraph(cell.Paragraphs[0]);
            cell._tableCellProperties!.TableCellWidth = new W.TableCellWidth {
                Type = W.TableWidthUnitValues.Pct, Width = column == 0 ? "3750" : "1250"
            };
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Docx })));
        var fonts = new OfficeIMO.Pdf.PdfCreateOptions();
        fonts.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, glyphAdvance, Enumerable.Range(32, 95).ToArray())));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false, PdfOptions = fonts }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters.ToArray();
        Assert.Equal(phrase ? "SASB SB" : "SASB", string.Concat(letters.Select(letter => letter.Value)));
        double first = letters[0].StartBaseLine.X;
        double second = letters.Where(letter => letter.Value == "S").ElementAt(1).StartBaseLine.X;
        Assert.InRange(second - first, expectedFirstColumnWidth - .001D, expectedFirstColumnWidth + .001D);
        Assert.Single(letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 3)).Distinct());
    }
}
