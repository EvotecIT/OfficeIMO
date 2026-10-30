using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableCellTextSizingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShrinkingReservesFixedInlineWidthInFlowAndCanvasTables(bool canvas) {
        PdfTableCell cell = new PdfTableCell(new[] {
            PdfTextRun.Inline(new PdfInlineBox(30, 10, background: global::OfficeIMO.OfficeColor.FromNormalizedRgb(.2D, .5D, .8D))),
            PdfTextRun.Normal("ABCDE")
        }).WithNoWrap().WithShrinkTextToFit();
        PdfTableStyle style = new PdfTableStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(12), HeaderRowCount = 0, HeaderBold = false,
            ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(50) }, MaxWidth = global::OfficeIMO.OfficeLength.Points(50), PreserveWidth = true,
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0), MinimumShrinkFontSize = global::OfficeIMO.OfficeLength.Points(4),
            FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(30) }, BorderColor = null
        };
        PdfDocument document = PdfDocument.Create(new PdfCreateOptions {
            DefaultFont = PdfStandardFont.Courier, PageWidth = global::OfficeIMO.OfficeLength.Points(200), PageHeight = global::OfficeIMO.OfficeLength.Points(150),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(20), MarginRight = global::OfficeIMO.OfficeLength.Points(20), MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20)
        });
        PdfTableCell[][] rows = new[] { new[] { cell } };
        if (canvas) document.Canvas(surface => surface.Table(rows, global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(30), style));
        else document.Table(rows, style: style);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        UglyToad.PdfPig.Content.Word word = Assert.Single(pdf.GetPage(1).GetWords());
        Assert.Equal("ABCDE", word.Text);
        Assert.All(word.Letters, letter => Assert.InRange(letter.PointSize, 6.66D, 6.67D));
        Assert.True(word.BoundingBox.Right <= 70.01D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DefaultCellShrinkMinimumAppliesToEveryExplicitRun(bool paragraphs) {
        var runs = new[] { new PdfTextRun("A", fontSize: global::OfficeIMO.OfficeLength.Points(30)), new PdfTextRun("B", fontSize: global::OfficeIMO.OfficeLength.Points(12)) };
        var cell = paragraphs ? new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs) }) : new PdfTableCell(runs);
        cell = cell.WithShrinkTextToFit();
        var style = new PdfTableStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(12), HeaderRowCount = 0, CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0),
            ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(4) }, MaxWidth = global::OfficeIMO.OfficeLength.Points(4), PreserveWidth = true, BorderColor = null
        };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(new PdfCreateOptions { DefaultFont = PdfStandardFont.Courier })
            .Table(new[] { new[] { cell } }, style: style).ToBytes());
        var letters = pdf.GetPages().SelectMany(x => x.Letters).ToArray();
        Assert.Equal("AB", string.Concat(letters.Select(x => x.Value)));
        Assert.All(letters, letter => Assert.InRange(letter.PointSize, 5.999D, 6.001D));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void CellTextShrinkingIsIndependentInFlowAndCanvasTables(bool canvas, bool explicitSize) {
        double targetSize = explicitSize ? 20 : 12;
        var target = explicitSize
            ? new PdfTableCell(new[] { new PdfTextRun("ShrinkTarget", fontSize: global::OfficeIMO.OfficeLength.Points(targetSize)) })
            : new PdfTableCell("ShrinkTarget");
        var rows = new[] { new[] {
            target.WithShrinkTextToFit().WithNoWrap().WithNamedDestination("Target"),
            new PdfTableCell("Neighbor"),
            target.WithNoWrap()
        } };
        var style = new PdfTableStyle {
            HeaderRowCount = 0, HeaderBold = false, FontSize = global::OfficeIMO.OfficeLength.Points(12),
            ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(90), global::OfficeIMO.OfficeLength.Points(50) },
            FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(40) },
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0), MinimumShrinkFontSize = global::OfficeIMO.OfficeLength.Points(4),
            PreserveWidth = true, MaxWidth = global::OfficeIMO.OfficeLength.Points(190), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
        };
        var document = PdfDocument.Create(new PdfCreateOptions {
            DefaultFont = PdfStandardFont.Courier, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12),
            PageWidth = global::OfficeIMO.OfficeLength.Points(300), PageHeight = global::OfficeIMO.OfficeLength.Points(200),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(20), MarginRight = global::OfficeIMO.OfficeLength.Points(20), MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20)
        });
        if (canvas) document.Canvas(surface => surface.Table(rows, global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(190), global::OfficeIMO.OfficeLength.Points(40), style));
        else document.Table(rows, style: style);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        var targets = words.Where(word => word.Text == "ShrinkTarget").OrderBy(word => word.BoundingBox.Left).ToArray();
        Assert.Equal(2, targets.Length);
        Assert.All(targets[0].Letters, letter => Assert.InRange(letter.PointSize, 4, 11));
        Assert.All(targets[1].Letters, letter => Assert.InRange(letter.PointSize, targetSize - .01, targetSize + .01));
        var neighbor = Assert.Single(words, word => word.Text == "Neighbor");
        Assert.All(neighbor.Letters, letter => Assert.InRange(letter.PointSize, 11.99, 12.01));
    }
}
