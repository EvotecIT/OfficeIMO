using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlRenderRequestTests {
    private const string TallHtml = "<style>html,body{margin:0}.block{height:700px;background:#336699}</style><div class='block'>Tall</div>";

    [Fact]
    public void BuiltInProfilesExposeStableUniqueAxisContracts() {
        Assert.Equal(6, HtmlRenderProfileContracts.All.Count);
        Assert.Equal(6, HtmlRenderProfileContracts.All.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count());

        HtmlRenderProfileContract print = HtmlRenderProfileContracts.Get(HtmlRenderIntentProfile.PrintPaged);
        Assert.Equal("print-paged-v1", print.Id);
        Assert.Equal(HtmlCssMediaContext.Print, print.CssMedia);
        Assert.Equal(HtmlRenderLayoutSurface.Paged, print.Surface);
        Assert.Equal(HtmlRenderPaginationPolicy.FragmentedReflow, print.Pagination);
        Assert.Equal(HtmlCapabilityCoverage.Qualified, print.Coverage);
        Assert.Equal(HtmlCapabilityPromotionState.StableDefault, print.Promotion);
        Assert.Contains(HtmlRenderEncoder.Pdf, print.Encoders);
        Assert.Contains(HtmlCapabilityEvidenceIds.H4PagedRepresentative, print.EvidenceIds);
        Assert.Contains(HtmlCapabilityEvidenceIds.H4PagedAdvancedHeldOut, print.EvidenceIds);

        HtmlRenderProfileContract screenSnapshotPaged = HtmlRenderProfileContracts.Get(HtmlRenderIntentProfile.ScreenSnapshotPaged);
        Assert.Equal(HtmlCapabilityCoverage.Qualified, screenSnapshotPaged.Coverage);
        Assert.Equal(HtmlCapabilityPromotionState.QualifiedOptIn, screenSnapshotPaged.Promotion);
        Assert.Contains(HtmlCapabilityEvidenceIds.H4PagedAdvancedHeldOut, screenSnapshotPaged.EvidenceIds);

        HtmlRenderProfileContract screenPaged = HtmlRenderProfileContracts.Get(HtmlRenderIntentProfile.ScreenMediaPaged);
        Assert.Equal(HtmlCssMediaContext.Screen, screenPaged.CssMedia);
        Assert.Equal(HtmlCapabilityCoverage.Unqualified, screenPaged.Coverage);
        Assert.Equal(HtmlCapabilityPromotionState.ExperimentalOptIn, screenPaged.Promotion);
    }

    [Fact]
    public void RequestSnapshotsOptionsAndKeepsEncoderIndependentFromLayout() {
        var source = new HtmlRenderOptions { ViewportWidth = 480D, ViewportHeight = 320D };
        HtmlRenderRequest request = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.Png, source,
            HtmlRenderDocumentState.EditedSnapshot);
        source.ViewportWidth = 999D;

        Assert.Equal(480D, request.Options.ViewportWidth);
        Assert.Equal(320D, request.Options.ViewportHeight);
        Assert.Equal(HtmlRenderDocumentState.EditedSnapshot, request.DocumentState);
        Assert.Equal(HtmlCssMediaContext.Screen, request.CssMedia);
        Assert.Equal(HtmlRenderLayoutSurface.Viewport, request.Surface);
        Assert.Equal(HtmlRenderPaginationPolicy.None, request.Pagination);
        Assert.Equal(HtmlRenderEncoder.Png, request.Encoder);

        HtmlRenderRequest svg = request.WithEncoder(HtmlRenderEncoder.Svg);
        Assert.Equal(request.ProfileId, svg.ProfileId);
        Assert.Equal(request.Surface, svg.Surface);
        Assert.Equal(request.Pagination, svg.Pagination);
        Assert.Empty(HtmlRenderProfileContracts.Validate());
    }

    [Fact]
    public void ScreenViewportClipsWhileScreenFullPageRetainsContentHeight() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 384D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(TallHtml);

        HtmlRenderResult viewport = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.DisplayList, options));
        HtmlRenderResult fullPage = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.DisplayList, options));

        Assert.Single(viewport.Document.Pages);
        Assert.Equal(120D, viewport.Document.Pages[0].Height);
        HtmlRenderClipGroup viewportClip = Assert.IsType<HtmlRenderClipGroup>(
            Assert.Single(viewport.Document.Pages[0].Scene));
        Assert.Equal(120D, viewportClip.ClipHeight);
        Assert.True(viewport.Surfaces[0].IsClipped);
        Assert.Equal(options.Scale, viewport.RequestedScale);
        Assert.Equal(options.BackgroundColor, viewport.BackgroundColor);
        Assert.Contains(HtmlCapabilityProviderIds.OfficeIMOHtml, viewport.DeclaredProviderIds);
        Assert.True(fullPage.Document.Pages[0].Height >= 700D);
        Assert.False(fullPage.Surfaces[0].IsClipped);
    }

    [Fact]
    public void ScreenFullPageAppliesBodyMarginPaddingAndWidthToTheRootFormattingBox() {
        const string html = "<body style='margin:8px;padding:12px;width:180px'><div style='height:10px;background:#336699'>Root box</div></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 320D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderResult result = HtmlRenderEngine.Execute(
            HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.DisplayList, options));

        HtmlRenderText text = Assert.Single(result.Document.Pages[0].Visuals.OfType<HtmlRenderText>());
        Assert.Equal(20D, text.X, 3);
        Assert.Equal(20D, text.Y, 3);
        HtmlRenderShape box = Assert.Single(result.Document.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => string.Equals(shape.Source, "div", StringComparison.Ordinal));
        Assert.Equal(20D, box.X, 3);
        Assert.Equal(180D, box.Width, 3);
    }

    [Fact]
    public void ScreenFullPageGrowsToHorizontalScrollContentWhileViewportRemainsBounded() {
        const string html = "<style>html,body{margin:0}</style><div style='width:500px;height:20px;background:#336699'>Wide</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 320D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        HtmlRenderResult fullPage = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.DisplayList, options));
        HtmlRenderResult viewport = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.DisplayList, options));

        Assert.Equal(500D, fullPage.Document.Pages[0].Width, 3);
        Assert.Equal(320D, viewport.Document.Pages[0].Width, 3);
        Assert.True(viewport.Surfaces[0].IsClipped);
    }

    [Fact]
    public void ScreenFullPageUsesDeclaredTableWidthAsItsBorderBoxAndIncludesBodyMargin() {
        const string html = "<body style='margin:8px'><table style='width:760px;border:2px solid #fff;background:#eee'><tr><td>Legacy</td></tr></table></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 640D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderResult result = HtmlRenderEngine.Execute(
            HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.DisplayList, options));

        Assert.Equal(768D, result.Document.Pages[0].Width, 3);
        HtmlRenderShape[] table = result.Document.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => string.Equals(shape.Source, "table", StringComparison.Ordinal))
            .ToArray();
        Assert.NotEmpty(table);
        Assert.All(table, shape => {
            Assert.Equal(8D, shape.X, 3);
            Assert.Equal(760D, shape.Width, 3);
        });
    }

    [Fact]
    public void PrintPagedClipsHorizontalBodyOverflowToThePageCanvas() {
        const string html = "<style>html,body{margin:0}</style><div style='width:500px;height:20px;background:#336699'>Wide PDF</div>";
        var options = new HtmlToPdfOptions {
            PageSize = new OfficePageSize(3D, 2D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false
        };

        byte[] pdf = HtmlConversionDocument.Parse(html).RenderToPdfBytes(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));

        OfficeIMO.Pdf.PdfDocumentInfo info = OfficeIMO.Pdf.PdfInspector.Inspect(pdf);
        Assert.Equal(1, info.PageCount);
        Assert.Equal(216D, info.Pages[0].Width, 2);
        Assert.Equal(144D, info.Pages[0].Height, 2);
    }

    [Fact]
    public void CssMediaIsIndependentFromPagedReflow() {
        const string html = "<style>.target{color:#0000ff}@media screen{.target{color:#ff0000}}@media print{.target{color:#008000}}</style><p class='target'>Media</p>";
        var options = new HtmlRenderOptions {
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false
        };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        HtmlRenderResult screen = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenMediaPaged, HtmlRenderEncoder.DisplayList, options));
        HtmlRenderResult print = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.DisplayList, options));

        HtmlRenderText screenText = Assert.Single(screen.Document.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Media");
        HtmlRenderText printText = Assert.Single(print.Document.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Media");
        Assert.Equal(OfficeColor.Red, screenText.Color);
        Assert.Equal(OfficeColor.FromRgb(0, 128, 0), printText.Color);
        Assert.Equal(HtmlRenderMode.Paged, screen.Document.Mode);
        Assert.Equal(HtmlRenderMode.Paged, print.Document.Mode);
    }

    [Fact]
    public void PageRulesUseTheRequestsCssMedia() {
        const string html = """
            <style media="screen">@page { size: 4in 2in; margin: 0 }</style>
            <style>@media print { @page { size: 4in 3in; margin: 0 } }</style>
            <p>Media page</p>
            """;
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        var options = new HtmlRenderOptions { HonorCssPageRules = true, Margins = HtmlRenderMargins.All(0D) };

        HtmlRenderResult screen = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenMediaPaged, options: options));
        HtmlRenderResult print = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: options));

        Assert.Equal(192D, screen.Document.Pages[0].Height);
        Assert.Equal(288D, print.Document.Pages[0].Height);
    }

    [Fact]
    public void RequestAxesCanBeOverriddenWithoutClaimingNamedProfileQualification() {
        HtmlRenderRequest custom = HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged)
            .WithCssMedia(HtmlCssMediaContext.Screen);

        Assert.Equal(HtmlCssMediaContext.Screen, custom.CssMedia);
        Assert.False(custom.MatchesNamedProfile);
        Assert.Equal(HtmlCapabilityCoverage.Unqualified, custom.Coverage);
        HtmlRenderRequest continuous = custom.WithAxes(
            HtmlCssMediaContext.Screen,
            HtmlRenderLayoutSurface.Continuous,
            HtmlRenderPaginationPolicy.None);
        Assert.Equal(HtmlRenderLayoutSurface.Continuous, continuous.Surface);
        Assert.Equal(HtmlRenderPaginationPolicy.None, continuous.Pagination);
        HtmlRenderResult result = HtmlRenderEngine.Execute(
            HtmlConversionDocument.Parse("<p>Custom axes</p>"), continuous);
        Assert.Equal(HtmlCapabilityCoverage.Unqualified, result.Coverage);
        Assert.Empty(result.DeclaredProviderIds);
        Assert.Throws<NotSupportedException>(() => custom.WithLayoutSurface(HtmlRenderLayoutSurface.Continuous));
        Assert.Throws<NotSupportedException>(() => custom.WithPagination(HtmlRenderPaginationPolicy.ElementAwarePlacement));
    }

    [Fact]
    public void ScreenSnapshotPagedSlicesOneCompletedScreenLayout() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 384D,
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false
        };
        HtmlRenderRequest request = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.DisplayList, options);

        HtmlRenderResult result = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(TallHtml), request);

        Assert.Equal(3, result.Document.Pages.Count);
        Assert.Equal(new[] { 0D, 288D, 576D }, result.Surfaces.Select(item => item.SourceOffsetY).ToArray());
        Assert.All(result.Document.Pages, page => {
            Assert.Equal(384D, page.Width);
            Assert.Equal(288D, page.Height);
        });
        Assert.All(result.Surfaces, surface => Assert.True(surface.IsClipped));
    }

    [Fact]
    public void SnapshotProjectionDoesNotRepeatLogicalTextAcrossSlices() {
        const string html = "<style>html,body{margin:0}p{height:280px;margin:0}</style>"
            + "<p>First marker</p><p>Second marker</p><p>Third marker</p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 384D,
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false
        };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderResult continuous = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, options: options));
        HtmlRenderRequest request = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, options);

        HtmlPdfRenderRequestResult result = source.RenderToPdfResult(request);
        string retainedText = result.RenderResult.Document.Text;
        string pdfText = OfficeIMO.Pdf.PdfReadDocument.Open(result.ToBytes()).ExtractText();

        foreach (string marker in new[] { "First marker", "Second marker", "Third marker" }) {
            int baselineCount = CountOccurrences(continuous.Document.Text, marker);
            Assert.True(baselineCount > 0);
            Assert.Equal(baselineCount, CountOccurrences(retainedText, marker));
            Assert.Equal(baselineCount, CountOccurrences(pdfText, marker));
        }
    }

    [Fact]
    public void SnapshotSelectionProjectsOnlyTheRequestedSliceWithinItsWorkBudget() {
        const string html = "<style>html,body{margin:0}p{height:280px;margin:0}</style>"
            + "<p>First</p><p>Second</p><p>Third</p>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 384D,
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false,
            MaxProjectedVisuals = 20
        };
        HtmlRenderRequest request = HtmlRenderRequest.Create(
                HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.DisplayList, options)
            .WithPageSet(HtmlRenderPageSet.Page(2));

        HtmlRenderResult result = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html), request);

        Assert.Single(result.Document.Pages);
        Assert.Equal(576D, result.Surfaces[0].SourceOffsetY);
        Assert.DoesNotContain("First", result.Document.Text, StringComparison.Ordinal);
        Assert.Contains("Third", result.Document.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ProjectionWorkAndStitchedSurfaceBoundsFailClosed() {
        var projectionOptions = new HtmlRenderOptions {
            ViewportWidth = 384D,
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false,
            MaxProjectedVisuals = 1
        };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(TallHtml);
        HtmlRenderRequest snapshot = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenSnapshotPaged, options: projectionOptions);
        Assert.Throws<InvalidOperationException>(() => HtmlRenderEngine.Execute(source, snapshot));

        var surfaceOptions = new HtmlRenderOptions {
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false,
            MaxSurfaceHeight = 500
        };
        HtmlRenderRequest stitched = HtmlRenderRequest.Create(
                HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, surfaceOptions)
            .WithPageSet(HtmlRenderPageSet.Stitched());
        Assert.Throws<InvalidOperationException>(() => HtmlRenderEngine.Execute(source, stitched));
    }

    [Fact]
    public void PageSelectionAndStitchingAreResolvedBeforeEncoding() {
        var options = new HtmlRenderOptions {
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false
        };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(TallHtml);

        HtmlRenderRequest selectedRequest = HtmlRenderRequest.Create(
                HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Svg, options)
            .WithPageSet(HtmlRenderPageSet.Page(1));
        HtmlRenderResult selected = HtmlRenderEngine.Execute(source, selectedRequest);
        Assert.Single(selected.Document.Pages);
        Assert.Equal(2, selected.Surfaces[0].SourcePageNumber);

        HtmlRenderRequest stitchedRequest = HtmlRenderRequest.Create(
                HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options)
            .WithPageSet(HtmlRenderPageSet.Stitched());
        HtmlRenderResult stitched = HtmlRenderEngine.Execute(source, stitchedRequest);
        Assert.Single(stitched.Document.Pages);
        Assert.Equal(HtmlRenderMode.Continuous, stitched.Document.Mode);
        Assert.Equal(864D, stitched.Document.Pages[0].Height);
        Assert.True(stitched.Surfaces[0].IsClipped);
        Assert.Equal(3, stitched.Surfaces[0].SourcePlacements.Count);
        Assert.Equal(new[] { 0D, 288D, 576D },
            stitched.Surfaces[0].SourcePlacements.Select(item => item.OutputOffsetY).ToArray());
    }

    [Fact]
    public void ExplicitImageRequestAndLegacyEntryPointUseTheSameSceneAndEncoder() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 320D,
            ViewportHeight = 180D,
            Margins = HtmlRenderMargins.All(0D)
        };
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<p style='color:#336699'>Shared output</p>");
        OfficeImageExportResult legacy = source.ExportImage(OfficeImageExportFormat.Png, options);

        HtmlRenderRequest request = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, options);
        OfficeImageExportResult explicitResult = Assert.Single(source.RenderImages(request));

        Assert.Equal(legacy.Width, explicitResult.Width);
        Assert.Equal(legacy.Height, explicitResult.Height);
        Assert.Equal(legacy.Bytes, explicitResult.Bytes);
    }

    [Fact]
    public void ExplicitPdfRequestCanSelectScreenMediaOrFrozenScreenSlicing() {
        const string html = "<style>html,body{margin:0}.content{height:700px}@media print{.content{height:100px}}</style><div class='content'>PDF</div>";
        var options = new HtmlToPdfOptions {
            PageSize = new OfficePageSize(4D, 3D),
            ViewportWidth = 384D,
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false
        };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        byte[] print = source.RenderToPdfBytes(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));
        byte[] screenPaged = source.RenderToPdfBytes(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenMediaPaged, HtmlRenderEncoder.Pdf, options));
        HtmlPdfRenderRequestResult snapshotResult = source.RenderToPdfResult(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, options));
        byte[] snapshotPaged = snapshotResult.ToBytes();

        Assert.Equal(1, OfficeIMO.Pdf.PdfInspector.Inspect(print).PageCount);
        Assert.Equal(3, OfficeIMO.Pdf.PdfInspector.Inspect(screenPaged).PageCount);
        Assert.Equal(3, OfficeIMO.Pdf.PdfInspector.Inspect(snapshotPaged).PageCount);
        Assert.Equal(HtmlRenderPaginationPolicy.FixedCanvasSlicing,
            snapshotResult.RenderResult.Request.Pagination);
        Assert.Equal(new[] { 0D, 288D, 576D },
            snapshotResult.RenderResult.Surfaces.Select(surface => surface.SourceOffsetY).ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DirectPdfByteEntryPointsKeepSerializationInsideTheRenderDeadline(bool useExplicitAsyncRequest) {
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<p>Serialization deadline</p>");
        // Keep this contract focused on serialization. The first process-wide system-font
        // discovery is part of rendering and may legitimately consume the short deadline.
        _ = source.ToPdfBytes();
        var timeout = TimeSpan.FromMilliseconds(100D);
        var provider = new SlowFirstEncryptionProvider(TimeSpan.FromMilliseconds(300D));
        var options = new HtmlToPdfOptions { RenderTimeout = timeout };
        options.PdfOptions.SetEncryption(new OfficeIMO.Pdf.PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = OfficeIMO.Pdf.PdfStandardEncryptionAlgorithm.Aes128,
            AesCryptographyProvider = provider
        });
        OfficeImageExportTimeoutException exception;
        if (useExplicitAsyncRequest) {
            HtmlRenderRequest request = HtmlRenderRequest.Create(
                HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options);
            exception = await Assert.ThrowsAsync<OfficeImageExportTimeoutException>(
                () => source.RenderToPdfBytesAsync(request));
        } else {
            exception = Assert.Throws<OfficeImageExportTimeoutException>(() => source.ToPdfBytes(options));
        }

        Assert.Equal(timeout, exception.Timeout);
        Assert.True(provider.EncryptOperations > 0);
    }

    [Fact]
    public void ProfileRejectsUndeclaredEncoder() {
        Assert.Throws<NotSupportedException>(() => HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ContinuousVector, HtmlRenderEncoder.Pdf));
    }

    private static int CountOccurrences(string value, string marker) {
        int count = 0;
        int offset = 0;
        while ((offset = value.IndexOf(marker, offset, StringComparison.Ordinal)) >= 0) {
            count++;
            offset += marker.Length;
        }
        return count;
    }

    private sealed class SlowFirstEncryptionProvider : OfficeIMO.Security.IOfficeAesCryptographyProvider {
        private readonly TimeSpan _delay;
        private int _encryptOperations;

        internal SlowFirstEncryptionProvider(TimeSpan delay) {
            _delay = delay;
        }

        public string Name => "HTML serialization deadline test AES";

        internal int EncryptOperations => _encryptOperations;

        public byte[] EncryptCbc(
            byte[] key,
            byte[] initializationVector,
            byte[] plaintext,
            OfficeIMO.Security.OfficeAesPadding padding) {
            if (Interlocked.Increment(ref _encryptOperations) == 1) {
                Thread.Sleep(_delay);
            }
            return OfficeIMO.Security.OfficeManagedAesCryptographyProvider.Default.EncryptCbc(
                key, initializationVector, plaintext, padding);
        }

        public byte[] DecryptCbc(
            byte[] key,
            byte[] initializationVector,
            byte[] ciphertext,
            OfficeIMO.Security.OfficeAesPadding padding) =>
            OfficeIMO.Security.OfficeManagedAesCryptographyProvider.Default.DecryptCbc(
                key, initializationVector, ciphertext, padding);
    }
}
