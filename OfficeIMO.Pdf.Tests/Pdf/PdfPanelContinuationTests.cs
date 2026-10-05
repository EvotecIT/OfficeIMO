using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPanelContinuationTests {
    [Theory]
    [InlineData("body")]
    [InlineData("columns")]
    [InlineData("row")]
    public void Panel_ContinuousDecorationScalesImageWithinItsFragmentInset(string layout) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, MaxGeneratedPages = 3
        };
        var panelStyle = new PdfPanelStyle {
            PaddingX = 0, PaddingY = 4, FragmentBottomInset = 4,
            SpacingBefore = 0, SpacingAfter = 0, RepeatFragmentDecoration = false
        };
        byte[] image = PdfPngTestImages.CreateRgbPng(2, 2);
        var imageStyle = new PdfImageStyle { ScaleDownToFit = true };
        void Compose(PdfContentBuilder content) => content.Panel(
            panel => panel.Image(image, 40, 160, style: imageStyle), panelStyle);
        PdfDocument document = PdfDocument.Create(options);
        if (layout == "columns") document.Columns(Compose, new PdfMultiColumnOptions { BalanceLastPage = false });
        else if (layout == "row") document.Row(row => row.RelativeColumn(Compose));
        else document.Panel(panel => panel.Image(image, 40, 160, style: imageStyle), panelStyle);

        byte[] bytes = document.ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        var placement = Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(152, placement.Height, 3);
        Assert.Equal(38, placement.Width, 3);
    }

    [Theory]
    [InlineData("body", true, 0D)]
    [InlineData("body", false, 0D)]
    [InlineData("body", false, 4D)]
    [InlineData("columns", true, 0D)]
    [InlineData("columns", false, 0D)]
    [InlineData("columns", false, 4D)]
    [InlineData("row", true, 0D)]
    [InlineData("row", false, 0D)]
    [InlineData("row", false, 4D)]
    public void Panel_ContinuationPreservesItsChosenPaddingPolicy(string layout, bool repeat, double inset) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFontSize = 12, MaxGeneratedPages = 6
        };
        var panelStyle = new PdfPanelStyle {
            PaddingX = 0, PaddingY = 4, SpacingBefore = 0, SpacingAfter = 0,
            BorderColor = PdfColor.Black, RepeatFragmentDecoration = repeat, FragmentBottomInset = inset
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
        bool reservesEveryFragment = repeat || inset > 0;
        int expectedPage = layout == "columns" ? (reservesEveryFragment ? 2 : 1) : (reservesEveryFragment ? 3 : 2);
        var marker = Assert.Single(Enumerable.Range(1, pdf.NumberOfPages)
            .SelectMany(page => pdf.GetPage(page).GetWords().Where(word => word.Text == "Marker015")
                .Select(word => (Page: page, Word: word))));
        Assert.Equal(expectedPage, marker.Page);
        var first = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Marker001");
        var continuation = Assert.Single(Enumerable.Range(1, pdf.NumberOfPages)
            .SelectMany(page => pdf.GetPage(page).GetWords()), word => word.Text == "Marker008");
        Assert.Equal(repeat ? 0D : 4D, continuation.BoundingBox.Top - first.BoundingBox.Top, 3);
        var markers = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords())
            .Where(word => word.Text.StartsWith("Marker", StringComparison.Ordinal)).Select(word => word.Text)
            .OrderBy(value => value, StringComparer.Ordinal);
        Assert.Equal(Enumerable.Range(1, 23).Select(index => $"Marker{index:D3}"), markers);
    }
}
