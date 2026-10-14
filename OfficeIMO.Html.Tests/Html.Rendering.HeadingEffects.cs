using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, false, false)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, false, true)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, false, false)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, false, true)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, true, false)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, true, true)]
    public void HtmlHeadings_TransformsPreserveNavigationCoordinatesAcrossProjection(
        HtmlRenderIntentProfile profile, bool stitched, bool clipped) {
        string html = "<style>html,body{height:200px;margin:0}h1{margin:0;font:16px/20px Arial}</style>"
            + "<a href='#target'>Go</a><h1 id='target' style='position:absolute;left:10px;top:50px;"
            + "width:80px;height:20px;transform:translate(2px,3px);transform-origin:0 0;"
            + (clipped ? "clip:rect(0,0,0,0);" : "") + "'>Heading</h1>";
        var options = new HtmlToPdfOptions {
            PageSize = new OfficePageSize(120D / 96D, 100D / 96D),
            Margins = HtmlRenderMargins.All(0), ViewportWidth = 120, ViewportHeight = 200,
            AutoFitWidePrintContent = false, HonorCssPageRules = false, AllowSystemFontFallback = false
        };
        HtmlRenderRequest request = HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options);
        if (stitched) request = request.WithPageSet(HtmlRenderPageSet.Stitched());
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderHeading heading = Assert.Single(HtmlRenderEngine.Execute(source, request).Document.Headings);
        byte[] bytes = source.RenderToPdfBytes(request);
        PdfCore.PdfOutlineItem outline = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).Outlines);
        PdfCore.PdfNamedDestination destination = Assert.Single(PdfCore.PdfDocumentReadResult.Load(bytes).NamedDestinations);

        Assert.Equal("Heading", heading.Text);
        Assert.Equal(1, heading.PageNumber);
        Assert.Equal(12D, heading.X, 6);
        Assert.Equal(53D, heading.Y, 6);
        Assert.Equal(1, outline.PageNumber);
        double height = PdfCore.PdfReadDocument.Open(bytes).Pages[0].GetPageSize().Height;
        Assert.Equal(height - 53D * 0.75D, outline.DestinationTop!.Value, 6);
        Assert.Equal(outline.DestinationTop.Value, destination.DestinationTop!.Value, 6);
        Assert.Equal(9D, destination.DestinationLeft!.Value, 6);
        if (clipped) Assert.DoesNotContain("Heading", PdfCore.PdfReadDocument.Open(bytes).ExtractText());
    }
}
