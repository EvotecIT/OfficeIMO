using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRichInlineLineHeightTests {
    [Theory]
    [InlineData("flow", false)]
    [InlineData("flow", true)]
    [InlineData("column", false)]
    [InlineData("column", true)]
    [InlineData("table", false)]
    [InlineData("table", true)]
    [InlineData("canvas", false)]
    [InlineData("canvas", true)]
    public void LoweredInlineBoxAndLargeTextShareTheReservedLineHeight(string frame, bool boxFirst) {
        byte[] bytes = Render(frame, boxFirst);
        Assert.Contains("0.2 0.5 0.8 rg", PdfEncoding.Latin1GetString(bytes), StringComparison.Ordinal);
        using var pdf = PdfPigDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var next = Assert.Single(letters, letter => letter.Value == "B");
        double boxBottom = first.StartBaseLine.Y - 20D;
        double nextLineTop = next.StartBaseLine.Y + 8D * .74D;
        Assert.True(boxBottom >= nextLineTop - .001D,
            $"Inline bottom {boxBottom} crossed the following line top {nextLineTop} in {frame}.");
        Assert.Equal(37.76D - 16D * .74D, first.StartBaseLine.Y - next.StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("table")]
    public void ExactSpacingRetainsItsAuthoredAdvanceWithLoweredInlineBoxes(string frame) {
        using var pdf = PdfPigDocument.Open(Render(frame, boxFirst: false, spacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(10))));
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var next = Assert.Single(letters, letter => letter.Value == "B");
        Assert.Equal(10D, first.StartBaseLine.Y - next.StartBaseLine.Y, 3);
    }

    [Fact]
    public void CompressedTextOnlySpacingRetainsItsAuthoredAdvance() {
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(8) })
            .Paragraph(p => p.FontSize(global::OfficeIMO.OfficeLength.Points(24)).Text("A\n").FontSize(global::OfficeIMO.OfficeLength.Points(8)).Text("B"),
                style: new PdfParagraphStyle { FontSize = global::OfficeIMO.OfficeLength.Points(8), LineSpacing = PdfLineSpacing.Multiple(.5) }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var first = Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A");
        var next = Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "B");
        Assert.Equal(12D - 16D * .74D, first.StartBaseLine.Y - next.StartBaseLine.Y, 3);
    }

    internal static byte[] Render(string frame, bool boxFirst, PdfLineSpacing? spacing = null) {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(300), PageHeight = global::OfficeIMO.OfficeLength.Points(300), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(8),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30), MarginBottom = global::OfficeIMO.OfficeLength.Points(30), MarginLeft = global::OfficeIMO.OfficeLength.Points(30), MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            CompressContentStreams = false };
        var text = new PdfTextRun("A", fontSize: global::OfficeIMO.OfficeLength.Points(24), font: PdfStandardFont.Helvetica);
        var box = PdfTextRun.Inline(new PdfInlineBox(global::OfficeIMO.OfficeLength.Points(20), global::OfficeIMO.OfficeLength.Points(32), background: global::OfficeIMO.OfficeColor.FromNormalizedRgb(.2, .5, .8), baselineOffset: global::OfficeIMO.OfficeLength.Points(-20)));
        var runs = boxFirst ? new[] { box, text, new PdfTextRun("\nB", fontSize: global::OfficeIMO.OfficeLength.Points(8)) }
            : new[] { text, box, new PdfTextRun("\nB", fontSize: global::OfficeIMO.OfficeLength.Points(8)) };
        var style = new PdfParagraphStyle { FontSize = global::OfficeIMO.OfficeLength.Points(8), LineSpacing = spacing ?? PdfLineSpacing.Multiple(1.4),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) };
        var document = PdfDocument.Create(options);
        if (frame == "flow") document.Paragraph(p => p.Runs(runs), style: style);
        else if (frame == "column") document.Content.Row(row => row.RelativeColumn(column => column.Paragraph(p => p.Runs(runs), style: style)));
        else if (frame == "canvas") document.Canvas(canvas => canvas.TextBox(runs, global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(240), global::OfficeIMO.OfficeLength.Points(240),
            new PdfCanvasTextBoxStyle { FontSize = global::OfficeIMO.OfficeLength.Points(8), LineHeight = global::OfficeIMO.OfficeLength.Points(11.2), PaddingX = global::OfficeIMO.OfficeLength.Points(0), PaddingY = global::OfficeIMO.OfficeLength.Points(0),
                Background = null, BorderColor = null }));
        else {
            var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, fontSize: global::OfficeIMO.OfficeLength.Points(8), lineSpacing: style.LineSpacing) });
            var tableStyle = PdfTableStyles.Minimal();
            tableStyle.HeaderRowCount = 0; tableStyle.FontSize = global::OfficeIMO.OfficeLength.Points(8);
            tableStyle.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0); tableStyle.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0); tableStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0);
            document.Table(new[] { new[] { cell } }, style: tableStyle);
        }
        return document.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }
}
