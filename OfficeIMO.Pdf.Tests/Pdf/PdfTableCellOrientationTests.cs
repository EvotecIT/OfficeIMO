using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableCellOrientationTests {
    [Theory]
    [InlineData("flow", -90)]
    [InlineData("column", -90)]
    [InlineData("canvas", -90)]
    [InlineData("flow", 90)]
    [InlineData("column", 90)]
    [InlineData("canvas", 90)]
    public void Turned_cell_text_survives_cell_copies_and_each_rendering_surface(string surface, int rotation) {
        var options = new PdfOptions { PageWidth = 240, PageHeight = 180,
            MarginTop = 24, MarginBottom = 24, MarginLeft = 24, MarginRight = 24,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 12 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { 100 };
        style.FixedRowHeights = new List<double?> { 80 };
        style.CellPaddingX = 4; style.CellPaddingY = 4;
        style.SpacingBefore = 0; style.SpacingAfter = 0;
        var cell = new PdfTableCell(new[] { PdfTextRun.Normal("ABCD") }, paragraphs: null, textRotation: rotation)
            .WithNoWrap().WithNamedDestination("turned-cell")
            .WithViewport(new PdfTableCellViewport(100, 80, 100, 80));
        var cells = new[] { new[] { cell } };
        PdfDocument document = PdfDocument.Create(options);
        byte[] bytes = surface switch {
            "flow" => document.Table(cells, style: style).ToBytes(),
            "column" => document.Row(row => row.PercentColumn(100, column => column.Table(cells, style: style))).ToBytes(),
            "canvas" => document.Canvas(canvas => canvas.Table(cells, 24, 24, 100, 80, style)).ToBytes(),
            _ => throw new ArgumentOutOfRangeException(nameof(surface))
        };
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal("ABCD", string.Concat(page.Letters.Select(letter => letter.Value)));
        Assert.All(page.Letters, letter => {
            Assert.Equal(0, letter.EndBaseLine.X - letter.StartBaseLine.X, 3);
            Assert.True((letter.EndBaseLine.Y - letter.StartBaseLine.Y) * rotation > 0);
            Assert.InRange(letter.StartBaseLine.X, 24, 124);
            Assert.InRange(letter.StartBaseLine.Y, 76, 156);
        });
    }
}
