using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableCellOrientationTests {
    [Theory]
    [InlineData(false, -90)]
    [InlineData(true, -90)]
    [InlineData(false, 90)]
    [InlineData(true, 90)]
    public void Automatic_columns_measure_turned_text_on_the_cross_axis(bool inColumn, int rotation) {
        var options = new PdfOptions { PageWidth = 260, PageHeight = 180,
            MarginLeft = 20, MarginRight = 20, MarginTop = 20, MarginBottom = 20,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 9 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0; style.AutoFitColumns = true;
        style.AutoFitWidthUsesContentMinimum = true;
        style.CellPaddingX = 2; style.CellPaddingY = 2;
        style.FixedRowHeights = new List<double?> { 100 };
        var cell = new PdfTableCell(new[] { PdfTextRun.Normal("ABCDEFGHIJKLMNOPQRSTUVWXYZ") },
            paragraphs: null, textRotation: rotation, linkUri: "https://example.com/turned-width");
        var cells = new[] { new[] { cell } };
        var document = PdfDocument.Create(options);
        if (inColumn) document.Row(row => row.PercentColumn(100, column => column.Table(cells, style: style)));
        else document.Table(cells, style: style);
        byte[] bytes = document.ToBytes();
        var link = Assert.Single(PdfInspector.Inspect(bytes).GetAnnotationsBySubtype("Link"));
        Assert.InRange(link.X2 - link.X1, 9, 20);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains(page.Letters, letter => letter.Value == "A");
        var firstLetter = page.Letters.First(letter => letter.Value == "A");
        Assert.InRange(firstLetter.StartBaseLine.X, link.X1, link.X2);
    }

    [Theory]
    [InlineData(false, -90)]
    [InlineData(true, -90)]
    [InlineData(false, 90)]
    [InlineData(true, 90)]
    public void Turned_content_is_consumed_once_when_a_horizontal_neighbor_continues(bool inColumn, int rotation) {
        var options = new PdfOptions { PageWidth = 260, PageHeight = 180,
            MarginLeft = 20, MarginRight = 20, MarginTop = 20, MarginBottom = 20,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 9 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 2; style.CellPaddingY = 2;
        style.ColumnWidthPoints = new List<double?> { 60, null };
        var turned = new PdfTableCell(new[] { PdfTextRun.Normal("TURNED") }, paragraphs: null,
            linkUri: "https://example.com/turned", textRotation: rotation,
            images: new[] { new PdfTableCellImage(PdfPngTestImages.CreateRgbPng(2, 2), 12, 12) });
        var cells = new[] { new[] { turned, new PdfTableCell(string.Join("\n",
            Enumerable.Range(1, 45).Select(line => "Line" + line.ToString("D2")))) } };
        var document = PdfDocument.Create(options);
        if (inColumn) document.Row(row => row.PercentColumn(100, column => column.Table(cells, style: style)));
        else document.Table(cells, style: style);
        byte[] bytes = document.ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages >= 3);
        string[] pageText = pdf.GetPages().Select(page => string.Concat(page.Letters.Select(letter => letter.Value))).ToArray();
        Assert.Contains("TURNED", pageText[0]);
        Assert.All(pageText.Skip(1), text => Assert.DoesNotContain("TURNED", text));
        Assert.Equal(1, pdf.GetPages().Sum(page => page.NumberOfImages));
        Assert.Single(PdfInspector.Inspect(bytes).GetAnnotationsBySubtype("Link"));
        string allText = string.Concat(pageText);
        for (int line = 1; line <= 45; line++) Assert.Contains("Line" + line.ToString("D2"), allText);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    public void Turned_authored_run_sizes_follow_the_same_shrink_policy_as_canvas(string surface) {
        double canvasSize = GetShrunkTurnedFontSize("canvas");
        Assert.True(canvasSize < 48);
        Assert.Equal(canvasSize, GetShrunkTurnedFontSize(surface), 3);
    }

    private static double GetShrunkTurnedFontSize(string surface) {
        var options = new PdfOptions { PageWidth = 240, PageHeight = 180,
            MarginTop = 24, MarginBottom = 24, MarginLeft = 24, MarginRight = 24 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0; style.FontSize = 12;
        style.ShrinkTextToFit = true;
        style.ColumnWidthPoints = new List<double?> { 100 };
        style.FixedRowHeights = new List<double?> { 80 };
        style.CellPaddingX = 4; style.CellPaddingY = 4;
        style.SpacingBefore = 0; style.SpacingAfter = 0;
        var cells = new[] { new[] { new PdfTableCell(new[] { new PdfTextRun("ABCD", fontSize: 48) },
            paragraphs: null, textRotation: -90) } };
        PdfDocument document = PdfDocument.Create(options);
        byte[] bytes = surface switch {
            "flow" => document.Table(cells, style: style).ToBytes(),
            "column" => document.Row(row => row.PercentColumn(100, column => column.Table(cells, style: style))).ToBytes(),
            "canvas" => document.Canvas(canvas => canvas.Table(cells, 24, 24, 100, 80, style)).ToBytes(),
            _ => throw new ArgumentOutOfRangeException(nameof(surface))
        };
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        return pdf.GetPage(1).Letters.First(letter => letter.Value == "A").FontSize;
    }

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
