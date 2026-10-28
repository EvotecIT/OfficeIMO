using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfTableParagraphLineBoxTests {
    [Theory]
    [InlineData(10D)]
    [InlineData(24D)]
    public void TableRowsUseTheirParagraphLineBoxesInsteadOfDocumentDefault(double defaultSize) {
        using WordDocument document = WordDocument.Create();
        var styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
            new FontSize { Val = (defaultSize * 2D).ToString(System.Globalization.CultureInfo.InvariantCulture) })));
        WordTable table = document.AddTable(2, 1);
        table._tableProperties!.TableStyle = null;
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 2000;
        table.GridColumnWidth = new List<int> { 2000 };
        for (int row = 0; row < 2; row++) {
            WordTableCell cell = table.Rows[row].Cells[0];
            cell.WidthType = WordTableWidthUnit.Dxa; cell.Width = 2000;
            cell.MarginTop = global::OfficeIMO.OfficeLength.Zero; cell.MarginBottom = global::OfficeIMO.OfficeLength.Zero;
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = row == 0 ? "A" : "B";
            paragraph.FontSize = 8;
            paragraph.FontFamily = "Arial";
            paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
            paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            paragraph.LineSpacingPoints = 10;
            paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
        }
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("AB", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(10D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
    }
}
