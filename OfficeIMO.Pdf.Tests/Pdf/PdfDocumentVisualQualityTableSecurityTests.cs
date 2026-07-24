using System.Collections.Generic;
using System.Linq;
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
                    TextRun.Normal("VisibleLine"),
                    TextRun.LineBreak(),
                    TextRun.Normal("ClippedSecret")
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
                    TextRun.Link(new string('W', 256), uri)
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
}
