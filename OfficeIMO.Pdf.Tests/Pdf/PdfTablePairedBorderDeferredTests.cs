using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData("uniform")]
    [InlineData("bottom-only")]
    [InlineData("top-only")]
    [InlineData("merged-bottom")]
    public void TablePairedBorders_DeferredBatchesPreservePaintAndTextGeometry(string mode) {
        byte[] whole = RenderDeferredPairedBorders(mode, 3), batched = RenderDeferredPairedBorders(mode, 1);
        AssertDeferredPairedOutput(whole, batched);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(10)]
    public void TablePairedBorders_DeferredBatchesPreserveCellSpacing(double spacing) =>
        AssertDeferredPairedOutput(RenderDeferredPairedBorders("bottom-only", 3, spacing), RenderDeferredPairedBorders("bottom-only", 1, spacing));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TablePairedBorders_DeferredPageTransitionsKeepPaintOnItsPage(bool header) =>
        AssertDeferredPairedOutput(RenderDeferredPairedBorders("uniform", 9, rowCount: 7, pageHeight: 120, header: header),
            RenderDeferredPairedBorders("uniform", 1, rowCount: 7, pageHeight: 120, header: header));

    private static void AssertDeferredPairedOutput(byte[] whole, byte[] batched) {
        using var expected = PdfPigDocument.Open(whole);
        using var actual = PdfPigDocument.Open(batched);
        Assert.Equal(expected.NumberOfPages, actual.NumberOfPages);
        for (int page = 1; page <= expected.NumberOfPages; page++) {
            var expectedWords = expected.GetPage(page).GetWords().ToArray();
            var actualWords = actual.GetPage(page).GetWords().ToArray();
            Assert.Equal(expectedWords.Length, actualWords.Length);
            for (int index = 0; index < expectedWords.Length; index++) {
                Assert.Equal(expectedWords[index].Text, actualWords[index].Text);
                Assert.Equal(expectedWords[index].BoundingBox.Left, actualWords[index].BoundingBox.Left, 3);
                Assert.Equal(expectedWords[index].BoundingBox.Top, actualWords[index].BoundingBox.Top, 3);
            }
            OfficeRasterImage expectedImage = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(whole, page));
            OfficeRasterImage actualImage = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(batched, page));
            for (int y = 0; y < expectedImage.Height; y++)
                for (int x = 0; x < expectedImage.Width; x++) Assert.Equal(expectedImage.GetPixel(x, y), actualImage.GetPixel(x, y));
        }
    }

    private static byte[] RenderDeferredPairedBorders(string mode, int batchSize, double spacing = 0, int rowCount = 3, double pageHeight = 220, bool header = false) {
        var style = MixedPairedBorderStyle();
        // Isolate border paint from the deferred minimum-row grouping contract.
        style.MinimumBodyRowsOnFirstPage = 0;
        style.MinimumBodyRowsOnLastPage = 0;
        style.CellSpacing = spacing;
        style.HeaderRowCount = header ? 1 : 0;
        style.FixedRowHeights = Enumerable.Repeat<double?>(20, rowCount).ToList();
        style.CellFills = new Dictionary<(int, int), PdfColor> {
            [(1, 0)] = PdfColor.FromRgb(255, 230, 160), [(1, 1)] = PdfColor.FromRgb(240, 240, 240)
        };
        if (mode == "uniform") {
            for (int row = 0; row < rowCount; row++)
                for (int column = 0; column < 2; column++) style.CellBorders![(row, column)] = new PdfCellBorder {
                    Color = PdfColor.FromRgb(255, 0, 0), Width = 2, LineStyle = PdfCellBorderLineStyle.TwoLine
                };
        } else if (mode == "top-only") {
            style.CellBorders![(1, 0)] = MixedPairedBorder(top: true);
        } else {
            style.CellBorders![(0, 0)] = MixedPairedBorder(bottom: true);
        }
        PdfTableCell[][] rows = Enumerable.Range(0, rowCount).Select(row => mode == "merged-bottom" && row == 0
            ? new[] { new PdfTableCell("gyp00A", columnSpan: 2) }
            : new[] { new PdfTableCell($"gyp{row:D2}A"), new PdfTableCell($"gyp{row:D2}B") }).ToArray();
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 320, PageHeight = pageHeight, MarginLeft = 30, MarginRight = 30, MarginTop = 30, MarginBottom = 30,
            CompressContentStreams = false
        });
        document.Compose(root => root.Page(page => page.Content(content => content.TableDeferred(() => rows, batchSize, style: style))));
        return document.ToBytes();
    }
}
