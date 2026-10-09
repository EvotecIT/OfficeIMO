using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("absolute", false)]
    [InlineData("absolute", true)]
    [InlineData("fixed", false)]
    [InlineData("fixed", true)]
    public void HtmlLegacyClip_EmptyPaintPreservesIncomingLinksToItsOwnTargets(string position, bool nested) {
        string content = nested ? "<div id='target'>Hidden</div>" : "Hidden";
        string html = "<style>body{margin:0}</style><a href='#target'>Go</a><div "
            + (nested ? "" : "id='target' ") + "style='position:" + position
            + ";left:10px;top:50px;width:40px;height:30px;background:red;clip:rect(1px,1px,1px,1px)'>"
            + content + "</div>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(source, LegacyClipOptions());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        byte[] pdf = source.ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Equal(new[] { "html-fragment:target" }, info.NamedDestinationNames);
        Assert.Equal(new[] { "html-fragment:target" }, info.LinkDestinationNames);
        Assert.Equal(OfficeColor.White, raster.GetPixel(11, 51));
        Assert.DoesNotContain("Hidden", string.Join("", PdfCore.PdfReadDocument.Open(pdf).Pages.Select(page => page.ExtractText())));
        Assert.Empty(rendered.Diagnostics);
    }

    [Theory]
    [InlineData("transform:translate(2px,3px)")]
    [InlineData("clip:rect(0,0,0,0);transform:translate(2px,3px)")]
    public void HtmlLegacyClip_EmptyAncestorPreservesNestedTransformedDestinations(string targetEffect) {
        string html = "<style>body{margin:0}</style><a href='#target'>Go</a>"
            + "<div style='position:absolute;left:10px;top:50px;width:40px;height:30px;clip:rect(0,0,0,0)'>"
            + "<div id='target' style='position:absolute;left:0;top:0;" + targetEffect + "'>Hidden</div></div>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(LegacyClipOptions()));
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Equal(new[] { "html-fragment:target" }, info.NamedDestinationNames);
        Assert.Equal(new[] { "html-fragment:target" }, info.LinkDestinationNames);
        PdfCore.PdfNamedDestination destination = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).NamedDestinations);
        Assert.Equal(1, destination.PageNumber);
        Assert.Equal(9D, destination.DestinationLeft!.Value, 6);
        // Viewport dimensions resolve CSS units; the default print paper remains A4.
        double paperHeight = PdfCore.PdfReadDocument.Open(pdf).Pages[0].GetPageSize().Height;
        Assert.Equal(paperHeight - 53D * 0.75D, destination.DestinationTop!.Value, 6);
        Assert.DoesNotContain("Hidden", string.Join("", PdfCore.PdfReadDocument.Open(pdf).Pages.Select(page => page.ExtractText())));
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, true, 0D, 1)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, false, 0D, 1)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, false, 150D, 3)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, true, 150D, 1)]
    public void HtmlLegacyClip_ProjectionPreservesClippedTargetsOnTheirTransformedPage(
        HtmlRenderIntentProfile profile, bool stitched, double shift, int expectedPage) {
        string html = "<style>html,body{height:300px;margin:0}</style><a href='#target'>Go</a>"
            + "<div style='position:absolute;left:10px;top:50px;width:40px;height:30px;clip:rect(0,0,0,0);"
            + "transform:translateY(" + shift.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "px);transform-origin:0 0'><div id='target'>Hidden</div></div>";
        var options = new HtmlToPdfOptions {
            PageSize = new OfficePageSize(300D / 96D, 100D / 96D),
            Margins = HtmlRenderMargins.All(0), ViewportWidth = 300, ViewportHeight = 300,
            HonorCssPageRules = false, AllowSystemFontFallback = false
        };
        HtmlRenderRequest request = HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options);
        if (stitched) request = request.WithPageSet(HtmlRenderPageSet.Stitched());
        byte[] pdf = HtmlConversionDocument.Parse(html).RenderToPdfBytes(request);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        PdfCore.PdfNamedDestination destination = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).NamedDestinations);

        Assert.Equal(new[] { "html-fragment:target" }, info.LinkDestinationNames);
        Assert.Equal(expectedPage, destination.PageNumber);
        Assert.Equal(7.5D, destination.DestinationLeft!.Value, 6);
        Assert.DoesNotContain("Hidden", PdfCore.PdfReadDocument.Open(pdf).ExtractText());
        double surfaceHeight = PdfCore.PdfReadDocument.Open(pdf).Pages[expectedPage - 1].GetPageSize().Height;
        double sourceY = 50D + shift - (stitched ? 0D : (expectedPage - 1) * 100D);
        Assert.Equal(surfaceHeight - sourceY * 0.75D, destination.DestinationTop!.Value, 6);
    }
}
