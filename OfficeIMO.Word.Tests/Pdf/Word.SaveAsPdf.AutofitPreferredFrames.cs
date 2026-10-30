using System.Linq;
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
        table.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(6000) * 635)));
        table.GridColumnWidth = new List<int> { 3000, 3000 }.Select(value => global::OfficeIMO.OfficeLength.Emus(checked((long)value * 635))).ToList();
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(3000) * 635)));
            cell.MarginLeft = cell.MarginRight = global::OfficeIMO.OfficeLength.Zero;
            cell.Paragraphs[0].Text = column == 0 ? "A" : "B";
        }
        var configuredStyle = new OfficeIMO.Pdf.PdfTableStyle {
            MaxWidth = global::OfficeIMO.OfficeLength.Points(100D), HeaderRowCount = 0, CellPaddingX = global::OfficeIMO.OfficeLength.Points(0D), BorderColor = null
        };
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = new OfficeIMO.Pdf.PdfCreateOptions { DefaultTableStyle = configuredStyle }
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        var left = Assert.Single(page.Letters, letter => letter.Value == "A");
        var right = Assert.Single(page.Letters, letter => letter.Value == "B");
        Assert.Equal(50D, right.StartBaseLine.X - left.StartBaseLine.X, 3);
        Assert.Equal(100D, configuredStyle.MaxWidth?.ToPoints());
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
        table.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)((int)(width * 20D)) * 635)));
        table.GridColumnWidth = new List<int> { 2250, 2250 }.Select(value => global::OfficeIMO.OfficeLength.Emus(checked((long)value * 635))).ToList();
        table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Nil, global::OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeColor.Black);
        if (spacing >= 0) table.StyleDetails.CellSpacing = global::OfficeIMO.OfficeLength.Points((spacing) / 20D);
        else table._tableProperties.TableCellSpacing = null;
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(2250) * 635)));
            cell.MarginLeft = cell.MarginRight = global::OfficeIMO.OfficeLength.Zero;
            cell.Paragraphs[0].Text = column == 0 ? "A" : "B";
            cell.Paragraphs[0].FontFamily = "Arial"; cell.Paragraphs[0].FontSize = global::OfficeIMO.OfficeLength.Points(12);
        }
        string xml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(325),OfficeIMO.OfficeLength.Points(400)), Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(72))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        var left = Assert.Single(page.Letters, letter => letter.Value == "A");
        var right = Assert.Single(page.Letters, letter => letter.Value == "B");
        Assert.Equal(width / 2D - Math.Max(0D, spacing) / 20D,
            right.StartBaseLine.X - left.StartBaseLine.X, 3);
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }
}
