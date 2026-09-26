using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, 500)]
    [InlineData(1400, 200)]
    public void SaveAsPdf_FullWidthFloatingTableMovesTextBelowOrToNextPage(int tableOffset, int pageHeight) {
        using var document = WordDocument.Create();
        var table = document.AddTable(1, 1);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 6400;
        table.Rows[0].Height = 900;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Floating";
        table._tableProperties!.TablePositionProperties = new TablePositionProperties {
            HorizontalAnchor = HorizontalAnchorValues.Margin, VerticalAnchor = VerticalAnchorValues.Text,
            TablePositionY = tableOffset, BottomFromText = 180
        };
        document.AddParagraph(string.Join(" ", Enumerable.Range(0, 40).Select(index => "word" + index)));
        var result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new PdfCore.PageSize(400, pageHeight), Margins = PdfCore.PageMargins.Uniform(40)
        });
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        var tableWord = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Floating");
        var followingWords = pdf.GetPage(1).GetWords().Where(word => word.Text.StartsWith("word")).ToList();
        Assert.All(followingWords, word => Assert.True(word.BoundingBox.Top <= pageHeight - 40 - tableOffset / 20D - 45 - 9 || word.BoundingBox.Bottom >= pageHeight - 40 - tableOffset / 20D));
        Assert.Equal(40, Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).Count(word => word.Text.StartsWith("word")));
        Assert.True(tableWord.BoundingBox.Top <= pageHeight - 40 - tableOffset / 20D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_FloatingTableWrapsTextAndRestoresWidthBelowTable(bool leftAligned) {
        using var document = WordDocument.Create();
        var table = document.AddTable(1, 1);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 2400;
        table.Rows[0].Height = 1600;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Floating";
        table._tableProperties!.TablePositionProperties = new TablePositionProperties {
            HorizontalAnchor = HorizontalAnchorValues.Margin,
            VerticalAnchor = VerticalAnchorValues.Text,
            TablePositionXAlignment = leftAligned ? HorizontalAlignmentValues.Left : HorizontalAlignmentValues.Right,
            LeftFromText = 180, RightFromText = 180,
            TablePositionY = -66
        };
        document.AddParagraph(string.Join(" ", Enumerable.Range(0, 100).Select(index => "word" + index)));
        var result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new PdfCore.PageSize(400, 500), Margins = PdfCore.PageMargins.Uniform(40)
        });
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        var words = pdf.GetPage(1).GetWords().ToList();
        var wrapped = words.Where(word => word.Text.StartsWith("word") && word.BoundingBox.Bottom > 380).ToList();
        Assert.NotEmpty(wrapped);
        if (leftAligned) Assert.All(wrapped, word => Assert.True(word.BoundingBox.Left >= 169));
        else Assert.All(wrapped, word => Assert.True(word.BoundingBox.Right <= 231));
        var below = words.Where(word => word.Text.StartsWith("word") && word.BoundingBox.Top < 370).ToList();
        Assert.NotEmpty(below);
        Assert.Contains(below, word => leftAligned ? word.BoundingBox.Left < 100 : word.BoundingBox.Right > 260);
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativePositionedTableWrapApproximation");
        Assert.Equal(100, Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).Count(word => word.Text.StartsWith("word")));
    }
}
