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
    public void PositionedTableDoesNotAdvanceFollowingFlow(bool rowColumn) {
        double FollowingTextY(bool consumesVerticalFlow) {
            var style = PdfTableStyles.Minimal();
            style.HeaderRowCount = 0;
            style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(36) };
            style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(80) };
            style.ConsumesVerticalFlow = consumesVerticalFlow;
            PdfDocument document = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(260), PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30), MarginRight = global::OfficeIMO.OfficeLength.Points(30), MarginTop = global::OfficeIMO.OfficeLength.Points(30), MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            });
            if (rowColumn) {
                document.Compose(compose => compose.Page(page => page.Content(content => content.Row(row =>
                    row.PercentColumn(100, column => {
                        column.Table(new[] { new[] { "Floating" } }, align: PdfAlign.Right, style: style);
                        column.Paragraph(paragraph => paragraph.Text("Following"));
                    })))));
            } else {
                document.Table(new[] { new[] { "Floating" } }, align: PdfAlign.Right, style: style);
                document.Paragraph(paragraph => paragraph.Text("Following"));
            }
            using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
            return pdf.GetPage(1).Letters.First(letter => letter.Value == "F" && letter.BoundingBox.Left < 50).BoundingBox.Bottom;
        }

        double flowingY = FollowingTextY(consumesVerticalFlow: true);
        double floatingY = FollowingTextY(consumesVerticalFlow: false);

        Assert.True(floatingY > flowingY + 25, $"Expected positioned table to preserve flow cursor; flowing={flowingY}, floating={floatingY}.");
    }

    [Fact]
    public void PositionedColumnTableThatSpansPagesKeepsFollowingTextBelowLastRow() {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ConsumesVerticalFlow = false;
        style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(38), global::OfficeIMO.OfficeLength.Points(38), global::OfficeIMO.OfficeLength.Points(38), global::OfficeIMO.OfficeLength.Points(38), global::OfficeIMO.OfficeLength.Points(38) };
        PdfDocument document = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260), PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30), MarginRight = global::OfficeIMO.OfficeLength.Points(30), MarginTop = global::OfficeIMO.OfficeLength.Points(30), MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
        });
        document.Compose(compose => compose.Page(page => page.Content(content => content.Row(row =>
            row.PercentColumn(100, column => {
                column.Table(new[] {
                    new[] { "R1" }, new[] { "R2" }, new[] { "R3" }, new[] { "R4" }, new[] { "R5" }
                }, style: style);
                column.Paragraph(paragraph => paragraph.Text("Following"));
            })))));

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.True(pdf.NumberOfPages > 1);
        var lastPageWords = pdf.GetPage(pdf.NumberOfPages).GetWords().ToList();
        double lastRowY = Assert.Single(lastPageWords, word => word.Text == "R5").BoundingBox.Bottom;
        double followingY = Assert.Single(lastPageWords, word => word.Text == "Following").BoundingBox.Bottom;
        Assert.True(followingY < lastRowY, $"Following text overlaps a continued positioned table: row={lastRowY}, text={followingY}.");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExactHeightTableRowsDoNotEmitClippedHiddenText(bool rowColumn) {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
        };
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(120) };
        style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(18) };
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
                            row.PercentColumn(100, column => column.Table(rows, style: style))))));
        } else {
            document.Table(rows, style: style);
        }

        byte[] bytes = document.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
        };
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(80) };
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
                            row.PercentColumn(100, column => column.Table(rows, style: style))))));
        } else {
            document.Table(rows, style: style);
        }

        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(document.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })).LinkAnnotations, annotation => annotation.Uri == uri);

        Assert.InRange(link.X1, 29.5D, 31D);
        Assert.InRange(link.X2, 30D, 112.5D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NoWrapTableParagraphTextDoesNotPaintAcrossAdjacentCells(bool rowColumn) {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12)
        };
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.BorderWidth = global::OfficeIMO.OfficeLength.Points(0D);
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0D);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0D);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(60), global::OfficeIMO.OfficeLength.Points(60) };
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
                            row.PercentColumn(100, column => column.Table(rows, style: style))))));
        } else {
            document.Table(rows, style: style);
        }

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(document.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })));
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

    [Fact]
    public void RightAlignedNoWrapTableParagraphRemainsVisibleInsideItsCell() {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(360),
            PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0);
        style.MaxWidth = global::OfficeIMO.OfficeLength.Points(200);
        style.PreserveWidth = true;
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(125), global::OfficeIMO.OfficeLength.Points(75) };
        PdfTextRun[] label = { PdfTextRun.Normal("Total Award Percent:") };
        PdfTableCell[][] rows = {
            new[] {
                new PdfTableCell(label, new[] { new PdfTableCellParagraph(label, align: PdfAlign.Right) }, noWrap: true),
                PdfTableCell.TextCell("6.00%")
            }
        };

        byte[] pdf = PdfDocument.Create(options).Table(rows, PdfAlign.Right, style).ToBytes();

        using PdfPigDocument parsed = PdfPigDocument.Open(pdf);
        Assert.Contains("Total Award Percent:", string.Concat(parsed.GetPages().Select(page => page.Text)), System.StringComparison.Ordinal);
    }
}
