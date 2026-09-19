using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRichFontUsageTests {
    [Theory]
    [InlineData(TableSurface.Flow)]
    [InlineData(TableSurface.RowColumn)]
    [InlineData(TableSurface.Canvas)]
    public void TableHeaderBoldCombinesWithRunItalicAcrossRenderingSurfaces(TableSurface surface) {
        var options = new PdfOptions {
            CompressContentStreams = false,
            PageWidth = 300D,
            PageHeight = 200D
        };
        var style = new PdfTableStyle {
            HeaderRowCount = 1,
            HeaderBold = true,
            ColumnWidthPoints = new List<double?> { 180D }
        };
        IReadOnlyList<PdfTableCell[]> rows = new[] {
            new[] {
                PdfTableCell.RichTextCell(new[] {
                    new PdfTextRun("StyledHeaderMarker", italic: true)
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
        string raw = Encoding.ASCII.GetString(bytes);

        Assert.True(bytes.AsSpan().StartsWith("%PDF"u8));
        Assert.Contains("/BaseFont /Helvetica-BoldOblique", raw, StringComparison.Ordinal);
        Assert.Contains("StyledHeaderMarker", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    public enum TableSurface {
        Flow,
        RowColumn,
        Canvas
    }
}
