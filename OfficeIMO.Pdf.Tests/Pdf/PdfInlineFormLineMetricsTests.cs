using OfficeIMO.Pdf;
using Xunit;
using Xunit.Abstractions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfInlineFormLineMetricsTests {
    private readonly ITestOutputHelper output;

    public PdfInlineFormLineMetricsTests(ITestOutputHelper output) => this.output = output;

    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("table", false)]
    [InlineData("canvas", false)]
    [InlineData("flow", true)]
    public void LargerSourceCheckBoxUsesTheSameLineAdvanceAndPaintAsItsOrdinaryRun(string frame, bool exactSpacing) {
        byte[] source = Render(frame, exactSpacing, field: true);
        Assert.True(Assert.Single(PdfReadDocument.Open(source).FormFields).IsCheckBox);
        byte[] flattened = PdfDocument.Load(source).Forms.Flatten().ToBytes();
        Assert.Empty(PdfReadDocument.Open(flattened).FormFields);

        using var actual = PdfPigDocument.Open(flattened);
        using var reference = PdfPigDocument.Open(Render(frame, exactSpacing, field: false));
        Assert.Equal(1, actual.NumberOfPages);
        Assert.Equal(reference.NumberOfPages, actual.NumberOfPages);
        var actualLetters = actual.GetPage(1).Letters;
        var referenceLetters = reference.GetPage(1).Letters;
        Assert.Equal(referenceLetters.Count, actualLetters.Count);
        foreach (string text in new[] { "P", "H", "B" }) {
            var expected = Assert.Single(referenceLetters, letter => letter.Value == text);
            var observed = Assert.Single(actualLetters, letter => letter.Value == text);
            Assert.Equal(expected.FontSize, observed.FontSize, 3);
            Assert.Equal(expected.StartBaseLine.X, observed.StartBaseLine.X, 3);
            Assert.Equal(expected.StartBaseLine.Y, observed.StartBaseLine.Y, 3);
            Assert.Equal(expected.BoundingBox.Left, observed.BoundingBox.Left, 3);
            Assert.Equal(expected.BoundingBox.Bottom, observed.BoundingBox.Bottom, 3);
            Assert.Equal(expected.BoundingBox.Right, observed.BoundingBox.Right, 3);
            Assert.Equal(expected.BoundingBox.Top, observed.BoundingBox.Top, 3);
        }

        var first = Assert.Single(actualLetters, letter => letter.Value == "H");
        var next = Assert.Single(actualLetters, letter => letter.Value == "B");
        Assert.Equal(20D, first.FontSize, 3);
        Assert.Equal(10D, next.FontSize, 3);
        double advance = first.StartBaseLine.Y - next.StartBaseLine.Y;
        if (exactSpacing) Assert.Equal(12D, advance, 3);
        else Assert.True(advance > 12D,
            "The larger source run must reserve more than the 10pt paragraph's default leading.");
        output.WriteLine($"{frame}, exact={exactSpacing}: source=20pt, paragraph=10pt, " +
            $"H baseline={first.StartBaseLine.Y:0.###}, B baseline={next.StartBaseLine.Y:0.###}, " +
            $"advance={advance:0.###}; cached H paint and all visible baselines match ordinary text.");
    }

    private static byte[] Render(string frame, bool exactSpacing, bool field) {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(300), PageHeight = global::OfficeIMO.OfficeLength.Points(300), DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30), MarginBottom = global::OfficeIMO.OfficeLength.Points(30), MarginLeft = global::OfficeIMO.OfficeLength.Points(30), MarginRight = global::OfficeIMO.OfficeLength.Points(30)
        };
        var source = PdfTextRun.Normal("H", fontSize: global::OfficeIMO.OfficeLength.Points(20), font: PdfStandardFont.Helvetica);
        PdfTextRun content = source;
        if (field) {
            var metrics = Assert.IsType<PdfWriter.PositionedTextBoxMeasurement>(
                PdfWriter.MeasurePositionedTextBox(source, options));
            var appearance = new PdfFormFieldStyle { BackgroundColor = null, BorderColor = null, BorderWidth = global::OfficeIMO.OfficeLength.Points(0) };
            var descriptor = new PdfInlineFormField(PdfCanvasFormFieldItem.CheckBox("LargerSource", true,
                "Yes", "Yes", 0, 0, metrics.Width, metrics.Ascent + metrics.Descent, appearance),
                source, PdfTextRun.Normal("O", fontSize: global::OfficeIMO.OfficeLength.Points(20), font: PdfStandardFont.Helvetica));
            var widget = new PdfInlineFormWidget(descriptor, metrics.Width, metrics.Ascent + metrics.Descent,
                -metrics.Descent, sourceRun: source);
            // This is the adapter's actual construction path; the source font size
            // is inherited by the inline wrapper before layout measures the line.
            content = PdfTextRun.Inline(widget);
        }
        var runs = new[] { PdfTextRun.Normal("P", fontSize: global::OfficeIMO.OfficeLength.Points(10)), content, PdfTextRun.Normal("\nB", fontSize: global::OfficeIMO.OfficeLength.Points(10)) };
        PdfLineSpacing? spacing = exactSpacing ? PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(12)) : null;
        var style = new PdfParagraphStyle { FontSize = global::OfficeIMO.OfficeLength.Points(10), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), LineSpacing = spacing };
        var document = PdfDocument.Create(options);
        if (frame == "flow") document.Paragraph(paragraph => paragraph.Runs(runs), style: style);
        else if (frame == "column") document.Content.Row(row => row.RelativeColumn(column =>
            column.Paragraph(paragraph => paragraph.Runs(runs), style: style)));
        else if (frame == "canvas") document.Canvas(canvas => canvas.TextBox(runs, global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(240), global::OfficeIMO.OfficeLength.Points(240),
            new PdfCanvasTextBoxStyle { FontSize = global::OfficeIMO.OfficeLength.Points(10), LineHeight = global::OfficeIMO.OfficeLength.Points(12), LineSpacing = spacing,
                PaddingX = global::OfficeIMO.OfficeLength.Points(0), PaddingY = global::OfficeIMO.OfficeLength.Points(0), Background = null, BorderColor = null }));
        else {
            var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, fontSize: global::OfficeIMO.OfficeLength.Points(10), lineSpacing: spacing) });
            var tableStyle = PdfTableStyles.Minimal();
            tableStyle.HeaderRowCount = 0; tableStyle.FontSize = global::OfficeIMO.OfficeLength.Points(10);
            tableStyle.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0); tableStyle.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0); tableStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0);
            document.Table(new[] { new[] { cell } }, style: tableStyle);
        }
        return document.ToBytes();
    }
}
