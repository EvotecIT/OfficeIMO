using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPanelInsetTests {
    [Theory]
    [InlineData(false, 18D, PdfAlign.Left, 38D)]
    [InlineData(true, 18D, PdfAlign.Left, 38D)]
    [InlineData(false, -10D, PdfAlign.Left, 10D)]
    [InlineData(false, 18D, PdfAlign.Right, 164D)]
    public void InsetsPlaceTheWholeFrameAndSurvivePageContinuation(bool row, double left, PdfAlign align, double expectedX) {
        using PdfDocument document = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = OfficeLength.Points(300), PageHeight = OfficeLength.Points(160), MarginLeft = OfficeLength.Points(20), MarginRight = OfficeLength.Points(20), MarginTop = OfficeLength.Points(20), MarginBottom = OfficeLength.Points(20)
        });
        var style = new PdfPanelStyle {
            LeftInset = OfficeLength.Points(left), RightInset = OfficeLength.Points(16),
            MaxWidth = align == PdfAlign.Right ? OfficeLength.Points(100) : null, Align = align,
            PaddingX = OfficeLength.Points(0), PaddingY = OfficeLength.Points(0), SpacingBefore = OfficeLength.Points(0), SpacingAfter = OfficeLength.Points(0),
            BorderColor = OfficeColor.Black, BorderWidth = OfficeLength.Points(1)
        };
        void Compose(PdfContentBuilder content) => content.Panel(panel => panel.Paragraph(
            p => p.Text(string.Join("\n", Enumerable.Range(1, 15).Select(i => $"Marker{i:D3}"))),
            style: new PdfParagraphStyle { LineSpacing = PdfLineSpacing.Exact(OfficeLength.Points(20)), SpacingAfter = OfficeLength.Points(0), WidowControl = false }), style);
        if (row) document.Row(columns => columns.RelativeColumn(Compose));
        else document.Panel(panel => Compose(panel), new PdfPanelStyle { PaddingX = OfficeLength.Points(0), PaddingY = OfficeLength.Points(0), BorderWidth = OfficeLength.Points(0) });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());

        Assert.True(pdf.NumberOfPages > 1);
        int count = 0;
        foreach (var page in pdf.GetPages()) {
            foreach (var word in page.GetWords()) {
                count++;
                Assert.Equal($"Marker{count:D3}", word.Text);
                Assert.Equal(expectedX, word.Letters[0].StartBaseLine.X, 2);
            }
            var frame = Assert.Single(page.Paths.Where(path => path.IsStroked)).GetBoundingRectangle()!.Value;
            Assert.Equal(expectedX, frame.Left, 2);
            Assert.Equal(264D, frame.Right, 2);
        }
        Assert.Equal(15, count);
    }
}
