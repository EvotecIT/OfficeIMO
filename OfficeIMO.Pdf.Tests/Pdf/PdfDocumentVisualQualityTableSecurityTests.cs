using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExactHeightTableRowsDoNotEmitClippedHiddenText(bool rowColumn) {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 0;
        style.CellPaddingY = 0;
        style.ColumnWidthPoints = new List<double?> { 120 };
        style.FixedRowHeights = new List<double?> { 18 };
        PdfTableCell[][] rows = {
            new[] {
                PdfTableCell.RichTextCell(new[] {
                    PdfTextRun.Normal("VisibleLine"),
                    PdfTextRun.LineBreak(),
                    PdfTextRun.Normal("ClippedSecret")
                })
            }
        };

        PdfDocument document = PdfDocument.Create(options);
        if (rowColumn) {
            document.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column.Table(rows, style: style))))));
        } else {
            document.Table(rows, style: style);
        }

        byte[] bytes = document.ToBytes();
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));

        Assert.Contains("VisibleLine", text);
        Assert.DoesNotContain("ClippedSecret", text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NoWrapTableRunLinksStayInsideCellClip(bool rowColumn) {
        const string uri = "https://example.com/no-wrap";
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 0;
        style.CellPaddingY = 0;
        style.ColumnWidthPoints = new List<double?> { 80 };
        PdfTableCell[][] rows = {
            new[] {
                PdfTableCell.RichTextCell(new[] {
                    PdfTextRun.Link(new string('W', 256), uri)
                }).WithNoWrap()
            }
        };

        PdfDocument document = PdfDocument.Create(options);
        if (rowColumn) {
            document.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column.Table(rows, style: style))))));
        } else {
            document.Table(rows, style: style);
        }

        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(document.ToBytes()).LinkAnnotations, annotation => annotation.Uri == uri);

        Assert.InRange(link.X1, 29.5D, 31D);
        Assert.InRange(link.X2, 30D, 112.5D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NoWrapTableParagraphTextDoesNotPaintAcrossAdjacentCells(bool rowColumn) {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 12
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.BorderWidth = 0D;
        style.CellPaddingX = 0D;
        style.CellPaddingY = 0D;
        style.ColumnWidthPoints = new List<double?> { 60, 60 };
        PdfTextRun[] runs = { PdfTextRun.Normal(new string('M', 80)) };
        PdfTableCell[][] rows = {
            new[] {
                new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs) }, noWrap: true),
                PdfTableCell.TextCell(string.Empty)
            }
        };

        PdfDocument document = PdfDocument.Create(options);
        if (rowColumn) {
            document.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column.Table(rows, style: style))))));
        } else {
            document.Table(rows, style: style);
        }

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(document.ToBytes()));
        bool paintedInAdjacentCell = false;
        for (int y = 28; y < 55 && !paintedInAdjacentCell; y++) {
            for (int x = 96; x < 145; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.A > 0 && pixel.R < 96 && pixel.G < 96 && pixel.B < 96) {
                    paintedInAdjacentCell = true;
                    break;
                }
            }
        }

        Assert.False(paintedInAdjacentCell, "No-wrap paragraph text must remain clipped before the adjacent table cell.");
    }
}
