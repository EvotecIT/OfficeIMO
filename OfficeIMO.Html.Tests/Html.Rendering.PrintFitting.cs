using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdf_PrintLayoutWidthFitsWideRowsOnPhysicalA4() {
        const string html = """
            <div style="width:66.666%">
              <div style="display:flex;flex-wrap:wrap">
                <div style="flex:none;width:310px;height:30px;background:#ff0000">FirstCard</div>
                <div style="flex:none;width:310px;height:30px;background:#0000ff"><a href="https://example.com/second">SecondCard</a></div>
              </div>
            </div>
            """;
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(24D),
            HonorCssPageRules = false,
            PrintLayoutWidthCssPixels = 1200D
        };

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(pdf);
        (double width, double height) = read.Pages[0].GetPageSize();
        PdfCore.PdfLogicalPage page = PdfCore.PdfDocumentReadResult.Load(pdf).Pages[0];
        PdfCore.PdfLogicalTextBlock first = Assert.Single(page.TextBlocks, block => block.Text.Contains("FirstCard", StringComparison.Ordinal));
        PdfCore.PdfLogicalTextBlock second = Assert.Single(page.TextBlocks, block => block.Text.Contains("SecondCard", StringComparison.Ordinal));
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri("https://example.com/second"));

        Assert.Equal(OfficePageSizes.A4.WidthInches * 72D, width, 2);
        Assert.Equal(OfficePageSizes.A4.HeightInches * 72D, height, 2);
        Assert.Equal(first.BaselineY, second.BaselineY, 2);
        Assert.InRange(first.XStart, 17D, 20D);
        Assert.True(second.XStart > first.XStart + 100D);
        Assert.InRange(link.SourceLink.X1, second.XStart - 5D, width);
        Assert.InRange(link.SourceLink.X2, link.SourceLink.X1, width);
        Assert.Contains("FirstCard", read.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("SecondCard", read.ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthRejectsConflictingPageRulesAndNonPrintIntent() {
        var options = new HtmlToPdfOptions {
            PrintLayoutWidthCssPixels = 1200D
        };
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<p>Content</p>");

        Assert.Throws<ArgumentException>(() => document.ToPdfBytes(options));

        options.HonorCssPageRules = false;
        HtmlRenderRequest screen = HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenMediaPaged, HtmlRenderEncoder.Pdf, options);
        Assert.Throws<ArgumentException>(() => document.RenderToPdfBytes(screen));
    }

    [Fact]
    public void HtmlPdf_PrintLayoutWidthKeepsInteractiveFieldAndFragmentTargetOnPhysicalPage() {
        const string html = """
            <a href="#target">Jump</a>
            <div style="height:200px"></div>
            <p id="target">Target</p>
            <input name="query" value="Hello" style="width:200px;height:30px">
            """;
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false,
            PrintLayoutWidthCssPixels = 1200D
        };

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfNamedDestination target = Assert.Single(info.NamedDestinations, destination => destination.Name == "html-fragment:target");
        PdfCore.PdfFormField field = Assert.Single(info.FormFields, item => item.Name == "query");

        Assert.InRange(target.DestinationTop!.Value, 650D, OfficePageSizes.A4.HeightInches * 72D);
        Assert.True(Assert.Single(field.Widgets).X2 <= OfficePageSizes.A4.WidthInches * 72D);
    }
}
