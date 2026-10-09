using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableHorizontalOffsetTests {
    [Theory]
    [InlineData("flow", PdfAlign.Left, -18)]
    [InlineData("flow", PdfAlign.Center, 12)]
    [InlineData("flow", PdfAlign.Right, -18)]
    [InlineData("columns", PdfAlign.Right, 12)]
    [InlineData("row", PdfAlign.Center, -18)]
    public void TranslationPreservesWidthsAndPagination(string frame, PdfAlign alignment, double offset) {
        using var baseline = PdfPigDocument.Open(Render(frame, alignment, 0));
        using var translated = PdfPigDocument.Open(Render(frame, alignment, offset));
        Assert.True(baseline.NumberOfPages >= 2);
        Assert.Equal(baseline.NumberOfPages, translated.NumberOfPages);
        for (int page = 1; page <= baseline.NumberOfPages; page++) {
            var original = baseline.GetPage(page).Letters;
            var moved = translated.GetPage(page).Letters;
            Assert.Equal(original.Count, moved.Count);
            for (int index = 0; index < original.Count; index++) {
                Assert.Equal(original[index].Value, moved[index].Value);
                Assert.InRange(Math.Abs(moved[index].StartBaseLine.X - original[index].StartBaseLine.X - offset), 0, .03);
                Assert.InRange(Math.Abs(moved[index].StartBaseLine.Y - original[index].StartBaseLine.Y), 0, .03);
            }
        }
    }

    [Fact]
    public void FloatingTranslationMovesTextReservationWithTheTable() {
        var style = new PdfTableStyle {
            HeaderRowCount = 0, MaxWidth = 120, PreserveWidth = true,
            CellPaddingX = 0, CellPaddingY = 0, MinRowHeight = 80,
            HorizontalOffset = 30, Position = new PdfTablePosition()
        };
        using var read = PdfPigDocument.Open(PdfDocument.Create(new PdfOptions {
            PageWidth = 400, PageHeight = 300, MarginLeft = 40, MarginRight = 40,
            MarginTop = 40, MarginBottom = 40
        }).Table(new[] { new[] { "floating" } }, style: style)
            .Paragraph(p => p.Text("alongside")).ToBytes());
        var words = read.GetPage(1).GetWords().ToArray();
        var table = Assert.Single(words, word => word.Text == "floating");
        var alongside = Assert.Single(words, word => word.Text == "alongside");
        Assert.InRange(table.Letters[0].StartBaseLine.X, 69.97, 70.03);
        Assert.True(alongside.Letters[0].StartBaseLine.X >= 190);
    }

    private static byte[] Render(string frame, PdfAlign alignment, double offset) {
        var style = new PdfTableStyle {
            LeftIndent = 6, HorizontalOffset = offset, MaxWidth = 120, PreserveWidth = true,
            ColumnWidthWeights = new List<double> { 1, 1 }, HeaderRowCount = 0,
            MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0,
            CellPaddingX = 0, CellPaddingY = 0, MinRowHeight = 50, FontSize = 9,
            BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0
        };
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 360, PageHeight = 240, MarginLeft = 30, MarginRight = 30,
            MarginTop = 30, MarginBottom = 30, DefaultFontSize = 9
        });
        string[][] rows = Enumerable.Range(1, 7).Select(index => new[] { "Left" + index, "Right" + index }).ToArray();
        if (frame == "columns") document.Columns(content => content.Table(rows, alignment, style),
            new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false });
        else if (frame == "row") document.Compose(builder => builder.Page(page => page.Content(content =>
            content.Row(row => row.PercentColumn(100, column => column.Table(rows, alignment, style))))));
        else document.Table(rows, alignment, style);
        return document.ToBytes();
    }
}
