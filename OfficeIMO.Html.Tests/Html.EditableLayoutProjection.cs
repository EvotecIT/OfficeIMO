using OfficeIMO.Html;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutProjectionTests {
    [Fact]
    public void ParagraphBreakStyledSpanAndSvgRegionsRemainInSemanticFlow() {
        const string html = "<div style='position:absolute;width:180px;height:50px'><p>Paragraph</p></div>" +
            "<div style='position:absolute;width:180px;height:50px'>Before<br>After</div>" +
            "<div style='position:absolute;width:180px;height:50px'><span style='color:red'>Styled</span></div>" +
            "<div style='position:absolute;width:180px;height:50px'><svg viewBox='0 0 10 10'><circle cx='5' cy='5' r='4'/></svg></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        string remaining = projection.RemainingDocument.DocumentElement!.OuterHtml;
        Assert.Contains("Paragraph", remaining, StringComparison.Ordinal);
        Assert.Contains("<br>", remaining, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("color", remaining, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<svg", remaining, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(4, projection.Diagnostics.Count(diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified));
    }

    [Fact]
    public void RendererLossDiagnosticsFlowThroughTheProjectionContract() {
        const string html = "<div style='position:absolute;width:180px;height:50px;" +
            "box-shadow:1px 1px 2px red,2px 2px 3px blue'>Limited effects</div>";
        var options = new HtmlRenderOptions { MaxBoxShadowLayers = 1 };

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html), options);

        Assert.Single(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.BoxShadowLayerLimit);
    }

    [Fact]
    public void CallerStylesheetsParticipateInRegionDiscoveryWithoutLeakingIntoSemanticFlow() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<div class='placed'>Caller styled</div>");
        var options = new HtmlRenderOptions();
        options.AdditionalStylesheets.Add(".placed{position:absolute;width:160px;height:40px}");

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            document, options);

        Assert.Single(projection.Regions);
        Assert.DoesNotContain("data-officeimo-render-stylesheet", projection.RemainingDocument.DocumentElement!.OuterHtml,
            StringComparison.Ordinal);
    }

    [Fact]
    public void AcceptedRegionsLeaveAPolicyNormalizedSemanticRemainder() {
        const string html = "<a href='javascript:alert(1)'>Unsafe</a>"
            + "<a href='docs/start.html'>Guide</a>"
            + "<img src='file:///secret/picture.png' alt='Rejected'>"
            + "<div style='position:absolute;width:160px;height:40px'>Projected</div>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            html,
            new HtmlConversionDocumentOptions {
                BaseUri = new Uri("https://example.test/root/page.html"),
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
                ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(document);

        Assert.Single(projection.Regions);
        string remaining = projection.RemainingDocument.DocumentElement!.OuterHtml;
        Assert.DoesNotContain("javascript:", remaining, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("file:///", remaining, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("https://example.test/root/docs/start.html", remaining, StringComparison.Ordinal);
    }

    [Fact]
    public void InheritedTypographyKeepsRegionInDiagnosedSemanticFlow() {
        const string html = "<body style='color:red;font-weight:700'>"
            + "<div style='position:absolute;width:160px;height:40px'>Inherited style</div></body>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Inherited style", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "inheritedTypography=true; semanticFlow=true");
    }

    [Fact]
    public void AuthoredTextAlignmentKeepsRegionInSemanticFlow() {
        const string html = "<div style='position:absolute;width:160px;height:40px;text-align:center'>Aligned</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("text-align:center", projection.RemainingDocument.Body!.InnerHtml,
            StringComparison.OrdinalIgnoreCase);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "semanticContent=true");
    }

    [Fact]
    public void MultiChildFlexAndGridRegionsStayInDiagnosedSemanticFlow() {
        const string html = "<div style='position:absolute;display:flex;gap:24px;width:240px;height:40px'>"
            + "Flex A<span>Flex B</span></div>"
            + "<div style='display:grid;grid-template-columns:1fr 1fr;width:240px;height:40px'>"
            + "<span>Grid A</span><span>Grid B</span></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Flex A", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains("Grid B", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Equal(2, projection.Diagnostics.Count(diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "multipleLayoutChildren=true; semanticFlow=true"));
    }

    [Fact]
    public void RegionTextUsesOnlyRenderedVisibleText() {
        const string html = "<div style='position:absolute;width:180px;height:50px'>Visible" +
            "<span style='display:none'>Hidden display</span>" +
            "<span style='visibility:hidden'>Hidden visibility</span></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
        Assert.Contains("Visible", region.SourceText, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden", region.SourceText, StringComparison.Ordinal);
    }

    [Fact]
    public void RichSemanticRegionsStayInNormalDocumentFlow() {
        const string html = "<div style='position:absolute;width:180px;height:50px'>" +
            "<strong>Bold</strong><a href='https://example.test'>Link</a></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Bold", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified);
    }

    [Fact]
    public void RegionRootBorderStaysInSemanticFlowInsteadOfBeingSilentlyFlattened() {
        const string html = "<div style='position:absolute;width:180px;height:50px;border:3px solid red'>Bordered</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Bordered", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "semanticContent=true");
    }

    [Fact]
    public void RegionAndAncestorPaintEffectsKeepContentInSemanticFlow() {
        const string html = "<div style='opacity:0'><div style='position:absolute;width:180px;height:50px'>Hidden by ancestor</div></div>" +
            "<div style='position:absolute;width:180px;height:50px;transform:rotate(4deg)'>Transformed</div>" +
            "<div style='position:absolute;width:180px;height:50px'>Visible<span style='opacity:0'>Hidden descendant</span></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Hidden by ancestor", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains("Transformed", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Equal(3, projection.Diagnostics.Count(diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported));
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Detail != null && diagnostic.Detail.Contains("opacity=0", StringComparison.Ordinal));
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Detail != null && diagnostic.Detail.Contains("transform=", StringComparison.Ordinal));
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Detail != null && diagnostic.Detail.Contains("descendant:opacity=0", StringComparison.Ordinal));
    }

    [Fact]
    public void MixedInlineTextAndPicturesKeepSourceOrderInSemanticFlow() {
        string image = "data:image/png;base64," + Convert.ToBase64String(
            OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:180px;height:50px'>Before" +
            "<img alt='Middle' src='" + image + "'>After</div>" +
            "<div style='position:absolute;width:180px;height:50px'><img alt='First' src='" + image + "'>After</div>" +
            "<div style='position:absolute;width:180px;height:50px'>Before<img alt='Last' src='" + image + "'></div>" +
            "<div style='position:absolute;width:180px;height:50px'><img alt='First' src='" + image + "'>Middle" +
            "<img alt='Last' src='" + image + "'></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.ProjectPreservingMixedInlineContent(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        string remaining = projection.RemainingDocument.Body!.InnerHtml;
        int before = remaining.IndexOf("Before", StringComparison.Ordinal);
        int picture = remaining.IndexOf("<img", StringComparison.OrdinalIgnoreCase);
        int after = remaining.IndexOf("After", StringComparison.Ordinal);
        Assert.True(before >= 0 && before < picture && picture < after);
        Assert.Equal(4, projection.Diagnostics.Count(diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "mixedInlinePictures=true"));
    }

    [Fact]
    public void PositionedFlexAndFloatingGridProduceSingleNativeRegions() {
        const string html = "<div style='position:absolute;display:flex;width:160px;height:40px'>Positioned flex</div>" +
            "<div style='float:right;display:grid;width:140px;height:40px'>Floating grid</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Equal(2, projection.Regions.Count);
        Assert.Contains(projection.Regions, region => region.RegionKind == HtmlRenderLayoutRegionKind.Positioned);
        Assert.Contains(projection.Regions, region => region.RegionKind == HtmlRenderLayoutRegionKind.Floating);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionFragmented);
    }

    [Fact]
    public void HiddenRegionImagesDoNotEnterNativeSourceAssociation() {
        string image = "data:image/png;base64," + Convert.ToBase64String(
            OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:160px;height:60px'>" +
            "<img alt='Hidden' src='" + image + "' style='display:none'>" +
            "<img alt='Visible' src='" + image + "' style='width:12px;height:12px'></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
        AngleSharp.Html.Dom.IHtmlImageElement source = Assert.Single(projection.GetSourceImages(region));
        Assert.Equal("Visible", source.AlternativeText);
    }

    [Fact]
    public void RegionImagesFollowRenderedFlexOrderByStableSourceMarker() {
        string image = "data:image/png;base64," + Convert.ToBase64String(
            OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;display:flex;width:160px;height:60px'>" +
            "<img alt='Dom first' src='" + image + "' style='order:2;width:12px;height:12px'>" +
            "<img alt='Visual first' src='" + image + "' style='order:1;width:12px;height:12px'></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
        IReadOnlyList<AngleSharp.Html.Dom.IHtmlImageElement> sources = projection.GetSourceImages(region);
        Assert.Equal(new[] { "Visual first", "Dom first" }, sources.Select(source => source.AlternativeText));
        IReadOnlyList<HtmlRenderImage> rendered = HtmlEditableLayoutProjector
            .EnumerateImages(region.Visuals, includeBackgroundImages: false)
            .Select(item => item.Image)
            .ToList();
        Assert.Equal(new[] { "img[officeimo-layout-image=2]", "img[officeimo-layout-image=1]" },
            rendered.Select(imageVisual => imageVisual.Source));
        Assert.Equal("Visual first", projection.GetSourceImage(rendered[0])!.AlternativeText);
        Assert.Equal("Dom first", projection.GetSourceImage(rendered[1])!.AlternativeText);
    }

    [Fact]
    public void RequestedPrintMediaUsesPagedRenderLayout() {
        const string html = "<style>.region{position:absolute;width:100px;height:20px}" +
            "@media print{.region{width:220px}}</style><div class='region'>Print region</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html),
            mediaContext: HtmlCssMediaContext.Print);

        Assert.Equal(HtmlRenderMode.Paged, projection.RenderedDocument.Mode);
        Assert.InRange(Assert.Single(projection.Regions).Width, 219.9D, 220.1D);
    }

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
            "<div class='flex'>Flex content</div>" +
            "<div class='grid'>Grid content</div></main>";

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
