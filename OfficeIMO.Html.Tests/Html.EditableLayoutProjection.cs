using OfficeIMO.Html;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutProjectionTests {
    [Fact]
    public void PublicVisualKindValuesRemainBackwardCompatible() {
        Assert.Equal(0, (int)HtmlRenderVisualKind.Shape);
        Assert.Equal(8, (int)HtmlRenderVisualKind.SemanticGroup);
        Assert.Equal(9, (int)HtmlRenderVisualKind.LogicalTextGroup);
        Assert.Equal(10, (int)HtmlRenderVisualKind.FormField);
        Assert.Equal(11, (int)HtmlRenderVisualKind.BookmarkAnchor);
        Assert.Equal(12, (int)HtmlRenderVisualKind.LayoutRegion);
    }

    [Fact]
    public void SourceAuthoredProjectorMarkerRemainsOrdinaryRenderableContent() {
        const string html = "<div data-officeimo-editable-layout-region='authored' style='display:grid'>Authored marker content</div>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(source, new HtmlRenderOptions());
        string pdfText = PdfCore.PdfReadDocument.Open(source.ToPdf()).ExtractText();

        Assert.Contains("Authored marker content", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Pages.SelectMany(page => page.Scene), visual =>
            visual.Kind == HtmlRenderVisualKind.LayoutRegion);
        Assert.Contains("Authored marker content", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void NestedCandidateRegionsProjectOnlyTheOwningOuterRegion() {
        const string html = "<section style='display:grid;width:260px'><div style='display:flex;width:180px'>Nested editable</div></section>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(HtmlConversionDocument.Parse(html));

        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
        Assert.Equal(HtmlRenderLayoutRegionKind.Grid, region.RegionKind);
        Assert.Equal("Nested editable", region.SourceText);
        Assert.Single(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
    }

    [Fact]
    public void PositionedFlexAndGridRegionsRetainEditableGeometryAndEffects() {
        const string html = "<style>" +
            ".absolute{position:absolute;left:40px;top:30px;width:200px;height:80px;z-index:7;background:#eef2ff;box-shadow:2px 3px 4px #444}" +
            ".flex{display:flex;width:240px;gap:8px}.grid{display:grid;grid-template-columns:1fr 1fr;width:220px}" +
            "</style><main><p>Normal flow</p>" +
            "<div class='absolute'>Positioned editable</div>" +
            "<div class='flex'><span>Flex A</span><span>Flex B</span></div>" +
            "<div class='grid'><span>Grid A</span><span>Grid B</span></div></main>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(HtmlConversionDocument.Parse(html));

        Assert.Equal(3, projection.Regions.Count);
        HtmlRenderLayoutRegion positioned = Assert.Single(projection.Regions, region =>
            region.RegionKind == HtmlRenderLayoutRegionKind.Positioned);
        Assert.Equal("Positioned editable", positioned.SourceText);
        Assert.Equal(7, positioned.ZIndex);
        Assert.Equal(1, positioned.BoxShadowLayerCount);
        Assert.NotNull(positioned.BackgroundColor);
        Assert.InRange(positioned.X, 87.9D, 88.1D);
        Assert.InRange(positioned.Y, 77.9D, 78.1D);
        Assert.Contains(projection.Regions, region => region.RegionKind == HtmlRenderLayoutRegionKind.Flex);
        Assert.Contains(projection.Regions, region => region.RegionKind == HtmlRenderLayoutRegionKind.Grid);
        Assert.All(projection.Regions, region => Assert.NotEmpty(region.Visuals));
        Assert.Equal(3, projection.Diagnostics.Count(diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected));
    }

    [Fact]
    public void FloatingRegionRetainsWrapSideAndNormalFlowStaysOutsideProjection() {
        const string html = "<p>Before</p><aside style='float:right;width:160px;background:#fff3cd'>Floating note</aside><p>After</p>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(HtmlConversionDocument.Parse(html));
        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);

        Assert.Equal(HtmlRenderLayoutRegionKind.Floating, region.RegionKind);
        Assert.Equal("right", region.FloatSide);
        Assert.Equal("Floating note", region.SourceText);
        Assert.Contains("Before", projection.RenderedDocument.Text, StringComparison.Ordinal);
        Assert.Contains("After", projection.RenderedDocument.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void MultiColumnProducerRegionKeepsNativePositionAndColumnFlowRemainsSemantic() {
        const string html = "<main style='column-count:2;column-gap:24px;width:420px'>" +
            "<p>Column one content that remains editable semantic flow.</p>" +
            "<aside style='float:right;width:120px;height:48px;background:#ddeeff'>Producer note</aside>" +
            "<p>Column two continuation that remains editable semantic flow.</p></main>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(HtmlConversionDocument.Parse(html));

        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
        Assert.Equal(HtmlRenderLayoutRegionKind.Floating, region.RegionKind);
        Assert.Equal("right", region.FloatSide);
        Assert.Equal("Producer note", region.SourceText);
        Assert.Contains("Column one content", projection.RenderedDocument.Text, StringComparison.Ordinal);
        Assert.Contains("Column two continuation", projection.RenderedDocument.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void RegionSplitAcrossColumnsOnOneSurfaceRemainsSemanticWithStableDiagnostic() {
        const string html = "<main style='column-count:2;column-fill:auto;height:120px;width:420px'>"
            + "<section style='display:flex;flex-direction:column;height:220px;width:180px'>"
            + "<span style='height:110px'>First fragment</span><span style='height:110px'>Second fragment</span>"
            + "</section></main>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("First fragment", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionFragmented
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("surfaces=1", StringComparison.Ordinal));
    }

    [Fact]
    public void PagedProducerRegionRetainsItsSurfaceAndFragmentedRegionGetsStableDiagnostic() {
        const string html = "<p>First page</p><section style='break-before:page'>" +
            "<div style='position:absolute;left:24px;top:32px;width:140px;height:50px'>Second-page anchor</div>" +
            "<div style='position:fixed;right:8px;top:8px;width:90px;height:24px'>Repeated header</div>" +
            "<div style='display:flex;width:220px'>" +
            "<span style='height:420px'>Tall A</span><span style='height:420px'>Tall B</span></div></section>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(12D)
        };

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html),
            options,
            HtmlCssMediaContext.Print);

        HtmlRenderLayoutRegion positioned = Assert.Single(projection.Regions, region =>
            region.RegionKind == HtmlRenderLayoutRegionKind.Positioned);
        Assert.InRange(positioned.SurfaceNumber, 1, projection.RenderedDocument.Pages.Count);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionFragmented);
        Assert.Contains("Repeated header", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
    }
}
