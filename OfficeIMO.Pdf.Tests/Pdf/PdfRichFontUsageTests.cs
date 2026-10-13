using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRichFontUsageTests {
    [Theory]
    [InlineData(TableSurface.Flow, false)]
    [InlineData(TableSurface.RowColumn, false)]
    [InlineData(TableSurface.Canvas, false)]
    [InlineData(TableSurface.Flow, true)]
    [InlineData(TableSurface.RowColumn, true)]
    [InlineData(TableSurface.Canvas, true)]
    public void TableHeaderBoldCombinesWithRunItalicAcrossRenderingSurfaces(TableSurface surface, bool namedFont) {
        var options = new PdfOptions {
            CompressContentStreams = false,
            PageWidth = 300D,
            PageHeight = 200D
        };
        if (namedFont) {
            byte[] font = ManagedTextShapingTestAssets.CreateFont("StyledHeaderMarker".Distinct().Select(ch => (int)ch).ToArray());
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("StyledHeader", font, font, font, font));
        }
        var style = new PdfTableStyle {
            HeaderRowCount = 1,
            HeaderBold = true,
            ColumnWidthPoints = new List<double?> { 180D }
        };
        IReadOnlyList<PdfTableCell[]> rows = new[] {
            new[] {
                PdfTableCell.RichTextCell(new[] {
                    new PdfTextRun("StyledHeaderMarker", italic: true, fontFamily: namedFont ? "StyledHeader" : null)
                })
            }
        };

        PdfDocument document = surface switch {
            TableSurface.Flow => PdfDocument.Create(options)
                .Table(rows, style: style),
            TableSurface.RowColumn => PdfDocument.Create(options)
                .Row(row => row.PercentColumn(100D, column => column.Table(rows, style: style))),
            TableSurface.Canvas => PdfDocument.Create(options)
                .Canvas(canvas => canvas.Table(rows, 40D, 40D, 180D, 60D, style)),
            _ => throw new ArgumentOutOfRangeException(nameof(surface))
        };

        byte[] bytes = document.ToBytes();
        Assert.True(bytes.AsSpan().StartsWith("%PDF"u8));
        Assert.Contains("StyledHeaderMarker", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = independent.GetPage(1).Letters;
        Assert.Equal("StyledHeaderMarker", string.Concat(letters.Select(letter => letter.Value)));
        string expectedFont = namedFont ? "StyledHeader-BoldItalic" : "Helvetica-BoldOblique";
        Assert.All(letters, letter => Assert.True(letter.FontName?.EndsWith(expectedFont, StringComparison.Ordinal) == true,
            "The painted header font must combine row bold with run italic; actual font: " + letter.FontName));
    }

    public enum TableSurface {
        Flow,
        RowColumn,
        Canvas
    }
}
