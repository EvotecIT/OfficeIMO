using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTableSignedIndentTests {
    [Theory]
    [InlineData("flow")]
    [InlineData("columns")]
    [InlineData("row")]
    public void SignedIndent_PreservesTableOriginAndColumnWidthsAcrossPageContinuation(string frame) {
        var options = new PdfOptions {
            PageWidth = 360, PageHeight = 240, MarginLeft = 30, MarginRight = 30,
            MarginTop = 30, MarginBottom = 30, DefaultFontSize = 9
        };
        var style = new PdfTableStyle {
            LeftIndent = -12, MaxWidth = 120, PreserveWidth = true,
            ColumnWidthWeights = new List<double> { 1, 1 }, HeaderRowCount = 0,
            MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0,
            CellPaddingX = 0, CellPaddingY = 0, MinRowHeight = 50, FontSize = 9,
            BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0
        };
        string[][] rows = Enumerable.Range(1, 7).Select(index =>
            new[] { "Left" + index, "Right" + index }).ToArray();
        PdfDocument document = PdfDocument.Create(options);
        if (frame == "columns") {
            document.Columns(content => content.Table(rows, style: style),
                new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false });
        } else if (frame == "row") {
            document.Compose(builder => builder.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style))))));
        } else {
            document.Table(rows, style: style);
        }

        using var read = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());
        Assert.True(read.NumberOfPages >= 2);
        var words = Enumerable.Range(1, read.NumberOfPages).SelectMany(page => read.GetPage(page).GetWords()).ToArray();
        for (int index = 1; index <= 7; index++) {
            var left = Assert.Single(words, word => word.Text == "Left" + index);
            var right = Assert.Single(words, word => word.Text == "Right" + index);
            double[] origins = frame == "columns" ? new[] { 18D, 178D } : new[] { 18D };
            Assert.InRange(origins.Min(origin => Math.Abs(left.Letters[0].StartBaseLine.X - origin)), 0D, .03D);
            Assert.InRange(right.Letters[0].StartBaseLine.X - left.Letters[0].StartBaseLine.X, 59.97D, 60.03D);
            Assert.InRange(Math.Abs(right.Letters[0].StartBaseLine.Y - left.Letters[0].StartBaseLine.Y), 0D, .03D);
        }
    }
}
