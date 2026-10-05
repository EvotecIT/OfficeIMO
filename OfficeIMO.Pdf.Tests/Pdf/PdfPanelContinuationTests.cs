using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPanelContinuationTests {
    [Theory]
    [InlineData("body", true)]
    [InlineData("body", false)]
    [InlineData("columns", true)]
    [InlineData("columns", false)]
    [InlineData("row", true)]
    [InlineData("row", false)]
    public void Panel_ContinuationPreservesItsChosenPaddingPolicy(string layout, bool repeat) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFontSize = 12, MaxGeneratedPages = 6
        };
        var panelStyle = new PdfPanelStyle {
            PaddingX = 0, PaddingY = 4, SpacingBefore = 0, SpacingAfter = 0,
            BorderColor = PdfColor.Black, RepeatFragmentDecoration = repeat
        };
        string text = string.Join("\n", Enumerable.Range(1, 23).Select(index => $"Marker{index:D3}"));
        void Compose(PdfContentBuilder content) => content.Panel(panel => panel.Paragraph(p => p.Text(text),
            style: new PdfParagraphStyle {
                LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
            }), panelStyle);
        PdfDocument document = PdfDocument.Create(options);
        if (layout == "columns") document.Columns(Compose, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false });
        else if (layout == "row") document.Row(row => row.RelativeColumn(Compose));
        else document.Panel(panel => panel.Paragraph(p => p.Text(text), style: new PdfParagraphStyle {
            LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
        }), panelStyle);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());
        int expectedPage = layout == "columns" ? (repeat ? 2 : 1) : (repeat ? 3 : 2);
        var marker = Assert.Single(Enumerable.Range(1, pdf.NumberOfPages)
            .SelectMany(page => pdf.GetPage(page).GetWords().Where(word => word.Text == "Marker015")
                .Select(word => (Page: page, Word: word))));
        Assert.Equal(expectedPage, marker.Page);
        var markers = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords())
            .Where(word => word.Text.StartsWith("Marker", StringComparison.Ordinal)).Select(word => word.Text)
            .OrderBy(value => value, StringComparer.Ordinal);
        Assert.Equal(Enumerable.Range(1, 23).Select(index => $"Marker{index:D3}"), markers);
    }
}
