using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData("flow")]
    [InlineData("row")]
    [InlineData("columns")]
    [InlineData("canvas")]
    public void TablePairedBorders_KeepStrokesClearOfCellText(string surface) {
        byte[] bytes = RenderPairedBorderTable(surface);
        var lines = PairedBorderLines(bytes);
        using var pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().Where(word => word.Text.StartsWith("gyp", StringComparison.Ordinal)).ToArray();
        Assert.Equal(4, words.Length);
        foreach (var word in words) {
            var box = word.BoundingBox;
            Assert.DoesNotContain(lines, line =>
                Math.Min(line.X1, line.X2) - 1 < box.Right && Math.Max(line.X1, line.X2) + 1 > box.Left &&
                Math.Min(line.Y1, line.Y2) - 1 < box.Top && Math.Max(line.Y1, line.Y2) + 1 > box.Bottom);
        }
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("row")]
    [InlineData("columns")]
    [InlineData("canvas")]
    public void TablePairedBorders_SharedEdgesHaveTwoVisibleTracks(string surface) {
        var lines = PairedBorderLines(RenderPairedBorderTable(surface));
        double[] vertical = lines.Where(line => Math.Abs(line.X1 - line.X2) < .001)
            .Select(line => Math.Round(line.X1, 3)).Distinct().OrderBy(x => x).ToArray();
        // Three boundaries, each with two tracks. Neighbouring cells must agree
        // on the shared pair instead of drawing inward on opposite sides.
        Assert.Equal(6, vertical.Length);
        double center = (vertical.First() + vertical.Last()) / 2;
        Assert.Equal(2, vertical.Count(x => Math.Abs(x - center) <= 4.001));
        Assert.Equal(4, vertical[3] - vertical[2], 3);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("row")]
    [InlineData("canvas")]
    public void TablePairedBorders_CellSpacingKeepsSeparateEdges(string surface) {
        var lines = PairedBorderLines(RenderPairedBorderTable(surface, spacing: 10));
        double[] vertical = lines.Where(line => Math.Abs(line.X1 - line.X2) < .001)
            .Select(line => Math.Round(line.X1, 3)).Distinct().ToArray();
        Assert.Equal(8, vertical.Length);
    }

    [Fact]
    public void TablePairedBorders_SquareCornersJoinBothTracks() {
        var lines = PairedBorderLines(RenderPairedBorderTable("flow"));
        foreach (var line in lines) {
            AssertJoined(line.X1, line.Y1, horizontal: line.Y1 == line.Y2);
            AssertJoined(line.X2, line.Y2, horizontal: line.Y1 == line.Y2);
        }
        void AssertJoined(double x, double y, bool horizontal) {
            Assert.Contains(lines, other => horizontal
                ? Math.Abs(other.X1 - other.X2) < .001 && Math.Abs(x - other.X1) < .001 &&
                    y >= Math.Min(other.Y1, other.Y2) - .001 && y <= Math.Max(other.Y1, other.Y2) + .001
                : Math.Abs(other.Y1 - other.Y2) < .001 && Math.Abs(y - other.Y1) < .001 &&
                    x >= Math.Min(other.X1, other.X2) - .001 && x <= Math.Max(other.X1, other.X2) + .001);
        }
    }

    [Fact]
    public void TablePairedBorders_ExistingPaddingKeepsTextPositions() {
        using var standard = PdfPigDocument.Open(RenderPairedBorderTable("flow", padding: 8, lineStyle: PdfCellBorderLineStyle.Standard));
        using var paired = PdfPigDocument.Open(RenderPairedBorderTable("flow", padding: 8));
        var expected = standard.GetPage(1).GetWords().ToArray();
        var actual = paired.GetPage(1).GetWords().ToArray();
        Assert.Equal(expected.Length, actual.Length);
        for (int index = 0; index < expected.Length; index++) {
            Assert.Equal(expected[index].Text, actual[index].Text);
            Assert.Equal(expected[index].BoundingBox.Left, actual[index].BoundingBox.Left, 3);
            Assert.Equal(expected[index].BoundingBox.Top, actual[index].BoundingBox.Top, 3);
        }
    }

    [Fact]
    public void TablePairedBorders_ContinuationKeepsEveryCellMarker() {
        using var pdf = PdfPigDocument.Open(RenderPairedBorderTable("fragment"));
        Assert.True(pdf.NumberOfPages > 1);
        var words = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).ToArray();
        for (int index = 0; index < 12; index++) Assert.Single(words, word => word.Text == $"gyp{index:D2}");
        foreach (string marker in new[] { "gypB", "gypC", "gypD" }) Assert.Single(words, word => word.Text == marker);
    }

    private static byte[] RenderPairedBorderTable(string surface, double spacing = 0, double padding = 0,
        PdfCellBorderLineStyle lineStyle = PdfCellBorderLineStyle.TwoLine) {
        var style = new PdfTableStyle {
            HeaderRowCount = 0, BorderColor = null, HeaderFill = null, RowStripeFill = null,
            CellPaddingX = padding, CellPaddingY = padding, FontSize = 12, LineHeight = 1,
            PreferredWidth = 200, CellSpacing = spacing,
            ColumnWidthPoints = new List<double?> { 100, 100 },
            CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder>()
        };
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 2; column++)
                style.CellBorders[(row, column)] = new PdfCellBorder {
                    Color = PdfColor.FromRgb(255, 0, 0), Width = 2, LineStyle = lineStyle
                };
        string[][] rows = { new[] { "gypA", "gypB" }, new[] { "gypC", "gypD" } };
        if (surface == "fragment") {
            style.AllowRowBreakAcrossPages = true;
            rows[0][0] = string.Join("\n", Enumerable.Range(0, 12).Select(index => $"gyp{index:D2}"));
        }
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = surface == "columns" ? 520 : 320, PageHeight = surface == "fragment" ? 110 : 220, MarginLeft = 30, MarginRight = 30,
            MarginTop = 30, MarginBottom = 30, CompressContentStreams = false
        });
        document.Compose(root => root.Page(page => page.Content(content => {
        switch (surface) {
            case "flow": content.Table(rows, style: style); break;
            case "fragment": content.Table(rows, style: style); break;
            case "row": content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style))); break;
            case "columns": content.Columns(inner => inner.Table(rows, style: style), new PdfMultiColumnOptions { ColumnCount = 2, BalanceLastPage = false }); break;
            case "canvas": content.Canvas(canvas => canvas.Table(rows, 30, 30, 200, 40, style)); break;
            default: throw new ArgumentException("Unknown surface", nameof(surface));
        }
        })));
        return document.ToBytes();
    }

    private static (double X1, double Y1, double X2, double Y2)[] PairedBorderLines(byte[] bytes) {
        string content = string.Join("\n", GetPageContentStreams(bytes, pageNumber: 1));
        const string number = @"(-?\d+(?:\.\d+)?)";
        return Regex.Matches(content, number + @"\s+" + number + @"\s+m\s+" + number + @"\s+" + number + @"\s+l\s+S")
            .Cast<Match>().Select(match => (
                double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture))).ToArray();
    }
}
