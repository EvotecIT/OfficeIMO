using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableTextReuseTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void RepeatedCellsRetainWrappingFontsAndPositions(bool rowColumn, bool embedded) {
        byte[]? fontData = embedded ? File.ReadAllBytes(PdfComplianceTestFonts.FindBundledTrueTypeFont()!) : null;
        byte[] implicitSize = CreateTable(rowColumn, fontData, explicitSize: false);
        byte[] explicitSize = CreateTable(rowColumn, fontData, explicitSize: true);
        AssertSameLetters(explicitSize, implicitSize);
    }

    [Fact]
    public void RepeatedCellsRetainStatefulLineBreakCallbacks() {
        int implicitCalls = 0, explicitCalls = 0;
        byte[] implicitSize = CreateTable(false, null, explicitSize: false, token => {
            implicitCalls++;
            return new[] { implicitCalls % 2 == 0 ? 5 : 4 };
        });
        byte[] explicitSize = CreateTable(false, null, explicitSize: true, token => {
            explicitCalls++;
            return new[] { explicitCalls % 2 == 0 ? 5 : 4 };
        });
        Assert.True(implicitCalls > 1);
        Assert.Equal(explicitCalls, implicitCalls);
        AssertSameLetters(explicitSize, implicitSize);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedViewportCellsRetainFullContentWidthsAndFragmentClipping(bool rowColumn) {
        AssertSameLetters(CreateViewportTable(rowColumn, explicitSize: true),
            CreateViewportTable(rowColumn, explicitSize: false));
    }

    private static byte[] CreateViewportTable(bool rowColumn, bool explicitSize) {
        var options = new PdfOptions {
            PageWidth = 400, PageHeight = 800,
            MarginLeft = 30, MarginRight = 30, MarginTop = 30, MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.FontSize = 10;
        style.CellPaddingX = 0;
        style.CellPaddingY = 0;
        style.ColumnWidthPoints = new List<double?> { 70, 70, 70 };
        style.FixedRowHeights = new List<double?> { 120, 120, 120 };
        var rows = new List<PdfTableCell[]>();
        for (int rowIndex = 0; rowIndex < 3; rowIndex++) {
            var cells = new PdfTableCell[3];
            for (int columnIndex = 0; columnIndex < cells.Length; columnIndex++) {
                double contentWidth = 70 * (columnIndex + 1);
                double offset = rowIndex < 2 ? 0 : Math.Min(70, contentWidth - 70);
                cells[columnIndex] = new PdfTableCell(new[] {
                    PdfTextRun.Normal("Alpha beta gamma delta epsilon zeta", fontSize: explicitSize ? 10 : null)
                }).WithViewport(new PdfTableCellViewport(contentWidth, 120, 70, 120, offsetX: offset));
            }
            rows.Add(cells);
        }
        PdfDocument document = PdfDocument.Create(options);
        if (rowColumn) {
            document.Compose(root => root.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style))))));
        } else {
            document.Table(rows, style: style);
        }
        return document.ToBytes();
    }

    private static byte[] CreateTable(bool rowColumn, byte[]? fontData, bool explicitSize, Func<string, IReadOnlyList<int>>? lineBreaks = null) {
        int count = rowColumn ? 12 : 96;
        var options = new PdfOptions {
            PageWidth = 400, PageHeight = 800,
            MarginLeft = 30, MarginRight = 30, MarginTop = 30, MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 10
        };
        if (fontData != null) options.EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "TableFont");
        if (lineBreaks != null) options.TextLineBreakCallback = lineBreaks;
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.FooterRowCount = 1;
        style.FontSize = 10;
        style.HeaderFontSize = 12;
        style.FooterFontSize = 8;
        style.ColumnWidthPoints = new List<double?> { 70, 190, 80 };
        var rows = new List<PdfTableCell[]>();
        for (int index = 0; index < count; index++) {
            double size = index == 0 ? 12 : index == count - 1 ? 8 : 10;
            string text = lineBreaks == null ? "Alpha beta gamma delta" : "alphaomegabetagamma";
            PdfTableCell Cell(string value, bool noWrap = false) =>
                new PdfTableCell(new[] { PdfTextRun.Normal(value, fontSize: explicitSize ? size : null) }).WithNoWrap(noWrap);
            rows.Add(new[] { Cell(text, index % 3 == 0), Cell(text), Cell("R" + index.ToString("D3")) });
        }
        PdfDocument document = PdfDocument.Create(options);
        if (rowColumn) {
            document.Compose(root => root.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style))))));
        } else {
            document.Table(rows, style: style);
        }
        return document.ToBytes();
    }

    private static void AssertSameLetters(byte[] expected, byte[] actual) {
        using var expectedPdf = UglyToad.PdfPig.PdfDocument.Open(expected);
        using var actualPdf = UglyToad.PdfPig.PdfDocument.Open(actual);
        Assert.Equal(expectedPdf.NumberOfPages, actualPdf.NumberOfPages);
        for (int page = 1; page <= expectedPdf.NumberOfPages; page++) {
            var expectedLetters = expectedPdf.GetPage(page).Letters;
            var actualLetters = actualPdf.GetPage(page).Letters;
            Assert.NotEmpty(actualLetters);
            Assert.Equal(expectedLetters.Select(letter => (letter.Value, letter.FontName, letter.PointSize, letter.GlyphRectangle)),
                actualLetters.Select(letter => (letter.Value, letter.FontName, letter.PointSize, letter.GlyphRectangle)));
        }
    }
}
