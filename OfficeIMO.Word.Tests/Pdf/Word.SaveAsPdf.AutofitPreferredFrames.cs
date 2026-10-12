using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_AutofitAuthoredFrameRetainsTheConfiguredDefaultTableMaximum() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2);
        table._tableProperties.TableStyle?.Remove();
        table.LayoutMode = WordTableLayoutMode.AutoFit;
        table.WidthType = WordTableWidthUnit.Dxa; table.Width = 6000;
        table.GridColumnWidth = new List<int> { 3000, 3000 };
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.WidthType = WordTableWidthUnit.Dxa; cell.Width = 3000;
            cell.MarginLeftWidth = cell.MarginRightWidth = 0;
            cell.Paragraphs[0].Text = column == 0 ? "A" : "B";
        }
        var configuredStyle = new OfficeIMO.Pdf.PdfTableStyle {
            MaxWidth = 100D, HeaderRowCount = 0, CellPaddingX = 0D, BorderColor = null
        };
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = new OfficeIMO.Pdf.PdfOptions { DefaultTableStyle = configuredStyle }
        }));
        var page = Assert.Single(pdf.GetPages());
        var left = Assert.Single(page.Letters, letter => letter.Value == "A");
        var right = Assert.Single(page.Letters, letter => letter.Value == "B");
        Assert.Equal(50D, right.StartBaseLine.X - left.StartBaseLine.X, 3);
        Assert.Equal(100D, configuredStyle.MaxWidth);
    }

    [Theory]
    [InlineData(100D, -1)]
    [InlineData(100D, 0)]
    [InlineData(100D, 120)]
    [InlineData(300D, -1)]
    [InlineData(300D, 0)]
    [InlineData(300D, 120)]
    public void SaveAsPdf_AutofitAbsoluteTableFrameIncludesItsSpacingAndCanExceedTheTextArea(double width, short spacing) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2, WordTableStyle.TableNormal);
        table.LayoutMode = WordTableLayoutMode.AutoFit;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = (int)(width * 20D);
        table.GridColumnWidth = new List<int> { 2250, 2250 };
        table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Nil, 0U, OfficeIMO.Drawing.OfficeColor.Black);
        if (spacing >= 0) table.StyleDetails.CellSpacing = spacing;
        else table._tableProperties.TableCellSpacing = null;
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.WidthType = WordTableWidthUnit.Dxa; cell.Width = 2250;
            cell.MarginLeftWidth = cell.MarginRightWidth = 0;
            cell.Paragraphs[0].Text = column == 0 ? "A" : "B";
            cell.Paragraphs[0].FontFamily = "Arial"; cell.Paragraphs[0].FontSize = 12;
        }
        string xml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new OfficeIMO.Pdf.PageSize(325, 400), Margins = OfficeIMO.Pdf.PageMargins.Uniform(72)
        }));
        var page = Assert.Single(pdf.GetPages());
        var left = Assert.Single(page.Letters, letter => letter.Value == "A");
        var right = Assert.Single(page.Letters, letter => letter.Value == "B");
        Assert.Equal(width / 2D - Math.Max(0D, spacing) / 20D,
            right.StartBaseLine.X - left.StartBaseLine.X, 3);
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }
}
