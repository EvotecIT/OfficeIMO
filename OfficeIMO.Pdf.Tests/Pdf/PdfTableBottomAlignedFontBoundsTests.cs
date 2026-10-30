using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableBottomAlignedFontBoundsTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void BottomAlignmentReservesNamedFontDescentWithoutChangingRowHeight(bool canvas, bool linked) {
        const string family = "Tall named table font";
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(240), PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24),
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(11) };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family,
            ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(1200, -400, 0, 400, 'A')));
        options.ReportDiagnosticsTo(report);
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FontSize = global::OfficeIMO.OfficeLength.Points(11); style.LineHeight = 1;
        style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(24) };
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(100) };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom };
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0); style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0);
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0); style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0);
        var cells = new[] { new[] { new PdfTableCell(new[] {
            new PdfTextRun("A", fontFamily: family, linkUri: linked ? "https://example.org/" : null)
        }) } };
        var document = PdfDocument.Create(options);
        byte[] bytes = canvas
            ? document.Canvas(c => c.Clip(24, 24, 100, 24, clipped => clipped.Table(cells, 24, 24, 100, 24, style))).ToBytes()
            : document.Table(cells, style: style).ToBytes();
        Assert.DoesNotContain(report.Warnings, warning => warning.Code.Contains("PaintClipped", StringComparison.Ordinal));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letter = Assert.Single(pdf.GetPage(1).Letters);
        Assert.InRange(letter.StartBaseLine.Y, 132D, 156D);
        Assert.Equal("A", letter.Value);
    }
}
