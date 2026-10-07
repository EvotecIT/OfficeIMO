using System.IO.Compression;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlRenderConsumerTests {
    private const string PagedHtml =
        "<style>html,body{margin:0}.page{height:280px;page-break-after:always}</style>" +
        "<div class='page'>One</div><div class='page'>Two</div><div class='page'>Three</div>";

    [Fact]
    public void OutputSurfaceMapsSelectedAndStitchedCoordinatesBackToTheirSources() {
        var options = new HtmlRenderOptions {
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false
        };
        HtmlConversionDocument source = HtmlConversionDocument.Parse(PagedHtml);

        HtmlRenderResult selected = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Geometry, options)
                .WithPageSet(HtmlRenderPageSet.Page(1)));
        Assert.True(selected.GetSurface(0).TryMapToSource(10D, 20D, out HtmlRenderSourcePoint? selectedPoint));
        Assert.Equal(2, selectedPoint!.SourcePageNumber);
        Assert.Equal(new HtmlRenderPoint(10D, 20D), selectedPoint.SourcePoint);

        HtmlRenderResult stitched = HtmlRenderEngine.Execute(source,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Geometry, options)
                .WithPageSet(HtmlRenderPageSet.Stitched()));
        Assert.True(stitched.GetSurface(0).TryMapToSource(10D, 300D, out HtmlRenderSourcePoint? stitchedPoint));
        Assert.Equal(2, stitchedPoint!.SourcePageNumber);
        Assert.Equal(new HtmlRenderPoint(10D, 12D), stitchedPoint.SourcePoint);
        Assert.False(stitched.GetSurface(0).TryMapToSource(stitched.GetSurface(0).Bounds.Right + 1D, 10D, out _));
    }

    [Fact]
    public void HitTestReturnsTopmostLinkAndRespectsViewportClip() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 240D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D)
        };
        var font = new OfficeFontInfo("Arial", 12D);
        var linkVisual = new HtmlRenderText("Top", 10D, 10D, 80D, 20D, font, OfficeColor.Black,
            OfficeTextAlignment.Left, 20D, 0, "https://example.test/top");
        var below = new HtmlRenderText("Below", 10D, 200D, 80D, 20D, font, OfficeColor.Black,
            OfficeTextAlignment.Left, 20D, 1, "https://example.test/below");
        var clip = new HtmlRenderClipGroup(0D, 0D, 240D, 120D, true, true,
            new HtmlRenderVisual[] { linkVisual, below }, 0);
        var page = new HtmlRenderPage(1, 240D, 120D, new[] { clip });
        var document = new HtmlRenderDocument(HtmlRenderMode.Continuous, new[] { page }, new HtmlDiagnosticReport());
        HtmlRenderRequest request = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.HitTest, options);
        var result = new HtmlRenderResult(request, document,
            new[] { new HtmlRenderSurfaceResult(0, 1, 240D, 120D, 0D, 0D, true) });
        HtmlRenderSurface surface = Assert.Single(result.OutputSurfaces);

        HtmlRenderHitTestReport top = surface.HitTest(
            linkVisual.X + Math.Min(1D, linkVisual.Width / 2D),
            linkVisual.Y + Math.Min(1D, linkVisual.Height / 2D),
            new HtmlRenderHitTestOptions {
            InteractiveOnly = true
        });
        HtmlRenderHitTestResult hit = Assert.Single(top.Matches);
        Assert.Contains("example.test/top", hit.Visual.LinkUri, StringComparison.Ordinal);
        Assert.True(hit.IsInteractive);
        Assert.Equal(HtmlRenderHitTestAccuracy.ClipAwareTransformedBounds, hit.Accuracy);
        Assert.NotNull(hit.SourcePoint);

        Assert.Empty(surface.HitTest(20D, 220D).Matches);
    }

    [Fact]
    public void HitTestReportsBoundedTraversal() {
        var html = new StringBuilder("<style>html,body{margin:0}span{display:block;height:20px}</style>");
        for (int index = 0; index < 20; index++) html.Append("<span>item</span>");
        HtmlRenderResult result = HtmlRenderEngine.Execute(
            HtmlConversionDocument.Parse(html.ToString()),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.HitTest));
        HtmlRenderVisual visual = result.Document.Pages[0].Visuals.First(item => item.Y >= 0D);

        HtmlRenderHitTestReport report = result.GetSurface(0).HitTest(
            visual.X + 1D,
            visual.Y + 1D,
            new HtmlRenderHitTestOptions { MaximumVisitedVisuals = 1 });

        Assert.True(report.IsTruncated);
        Assert.True(report.VisitedVisualCount > 1);
    }

    [Fact]
    public void HitTestAppliesAffineTransformsAndFreeformClips() {
        var font = new OfficeFontInfo("Arial", 12D);
        var link = new HtmlRenderText("Target", 0D, 0D, 100D, 100D, font, OfficeColor.Black,
            OfficeTextAlignment.Left, 20D, 0, "https://example.test/target");
        OfficeClipPath triangle = OfficeClipPath.Path(
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(100D, 0D),
            OfficePathCommand.LineTo(0D, 100D),
            OfficePathCommand.Close());
        var clipped = new HtmlRenderPathClipGroup(0D, 0D, triangle, new[] { link }, 0);
        var transformed = new HtmlRenderEffectGroup(50D, 30D, 100D, 100D,
            OfficeTransform.Translate(50D, 30D), 1D, new[] { clipped }, 0, "translated triangle");
        var page = new HtmlRenderPage(1, 200D, 180D, new[] { transformed });
        var document = new HtmlRenderDocument(HtmlRenderMode.Continuous, new[] { page }, new HtmlDiagnosticReport());
        var options = new HtmlRenderOptions { ViewportWidth = 200D, ViewportHeight = 180D };
        var result = new HtmlRenderResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.HitTest, options),
            document,
            new[] { new HtmlRenderSurfaceResult(0, 1, 200D, 180D, 0D, 0D, false) });

        HtmlRenderHitTestResult inside = Assert.Single(result.GetSurface(0).HitTest(
            60D, 40D, new HtmlRenderHitTestOptions { InteractiveOnly = true }).Matches);
        Assert.Equal(HtmlRenderHitTestAccuracy.ClipAwareTransformedBounds, inside.Accuracy);
        Assert.Empty(result.GetSurface(0).HitTest(
            140D, 120D, new HtmlRenderHitTestOptions { InteractiveOnly = true }).Matches);
        HtmlRenderHitTestReport bounded = result.GetSurface(0).HitTest(
            60D, 40D, new HtmlRenderHitTestOptions { InteractiveOnly = true, MaximumPathSegments = 1 });
        Assert.True(bounded.IsTruncated);
        Assert.Empty(bounded.Matches);
    }

    [Fact]
    public void HitTestRestoresContourStartBeforeCurveAfterClose() {
        var link = new HtmlRenderText("Target", 0D, 0D, 100D, 100D,
            new OfficeFontInfo("Arial", 12D), OfficeColor.Black,
            OfficeTextAlignment.Left, 20D, 0, "https://example.test/target");
        OfficeClipPath path = OfficeClipPath.Path(
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(40D, 0D),
            OfficePathCommand.LineTo(40D, 40D),
            OfficePathCommand.Close(),
            OfficePathCommand.QuadraticBezierTo(0D, 100D, 100D, 100D),
            OfficePathCommand.LineTo(100D, 0D),
            OfficePathCommand.Close());
        var clipped = new HtmlRenderPathClipGroup(0D, 0D, path, new[] { link }, 0);
        var page = new HtmlRenderPage(1, 100D, 100D, new[] { clipped });
        var document = new HtmlRenderDocument(HtmlRenderMode.Continuous, new[] { page }, new HtmlDiagnosticReport());
        var result = new HtmlRenderResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.HitTest,
                new HtmlRenderOptions { ViewportWidth = 100D, ViewportHeight = 100D }),
            document,
            new[] { new HtmlRenderSurfaceResult(0, 1, 100D, 100D, 0D, 0D, false) });

        HtmlRenderHitTestReport report = result.GetSurface(0).HitTest(
            5D, 10D, new HtmlRenderHitTestOptions { InteractiveOnly = true });

        Assert.False(report.IsTruncated);
        Assert.Single(report.Matches);
    }

    [Fact]
    public void HitTestDoesNotChargeARedundantCloseAgainstTheExactPathLimit() {
        var link = new HtmlRenderText("Target", 0D, 0D, 100D, 100D,
            new OfficeFontInfo("Arial", 12D), OfficeColor.Black,
            OfficeTextAlignment.Left, 20D, 0, "https://example.test/target");
        OfficeClipPath path = OfficeClipPath.Path(
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(40D, 0D),
            OfficePathCommand.LineTo(0D, 0D),
            OfficePathCommand.Close(),
            OfficePathCommand.QuadraticBezierTo(0D, 100D, 100D, 100D),
            OfficePathCommand.LineTo(100D, 0D),
            OfficePathCommand.Close());
        var clipped = new HtmlRenderPathClipGroup(0D, 0D, path, new[] { link }, 0);
        var page = new HtmlRenderPage(1, 100D, 100D, new[] { clipped });
        var document = new HtmlRenderDocument(HtmlRenderMode.Continuous, new[] { page }, new HtmlDiagnosticReport());
        var result = new HtmlRenderResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.HitTest,
                new HtmlRenderOptions { ViewportWidth = 100D, ViewportHeight = 100D }),
            document,
            new[] { new HtmlRenderSurfaceResult(0, 1, 100D, 100D, 0D, 0D, false) });

        HtmlRenderHitTestReport report = result.GetSurface(0).HitTest(
            50D, 40D, new HtmlRenderHitTestOptions {
                InteractiveOnly = true,
                MaximumPathSegments = 20
            });

        Assert.False(report.IsTruncated);
        Assert.Single(report.Matches);
    }

    [Fact]
    public void RetainedShapeSnapshotsCannotMutateLaterConsumers() {
        OfficeShape shape = OfficeShape.Rectangle(60D, 30D);
        shape.FillColor = OfficeColor.FromRgb(0x33, 0x66, 0x99);
        var visual = new HtmlRenderShape(shape, 10D, 10D, 0);
        var page = new HtmlRenderPage(1, 100D, 60D, new[] { visual });
        var document = new HtmlRenderDocument(HtmlRenderMode.Continuous, new[] { page }, new HtmlDiagnosticReport());
        var result = new HtmlRenderResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.Svg,
                new HtmlRenderOptions { ViewportWidth = 100D, ViewportHeight = 60D }),
            document,
            new[] { new HtmlRenderSurfaceResult(0, 1, 100D, 60D, 0D, 0D, false) });
        byte[] before = result.ExportArchive().Bytes;

        OfficeShape detached = visual.Shape;
        detached.Width = 1D;
        detached.FillColor = OfficeColor.Red;

        Assert.Equal(60D, visual.Shape.Width);
        Assert.Equal(OfficeColor.FromRgb(0x33, 0x66, 0x99), visual.Shape.FillColor);
        Assert.Equal(before, result.ExportArchive().Bytes);
    }

    [Fact]
    public void SvgArchiveIsDeterministicAndPreservesSelectedSurfaceEvidence() {
        var options = new HtmlRenderOptions {
            PageSize = new OfficePageSize(4D, 3D),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false,
            BackgroundColor = OfficeColor.FromRgba(10, 20, 30, 200)
        };
        HtmlRenderResult result = HtmlRenderEngine.Execute(
            HtmlConversionDocument.Parse(PagedHtml),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Svg, options)
                .WithPageSet(HtmlRenderPageSet.Pages(1, 2)));

        HtmlRenderArchiveResult first = result.ExportArchive();
        HtmlRenderArchiveResult second = result.ExportArchive();

        Assert.Equal(first.Bytes, second.Bytes);
        Assert.Equal(2, first.Manifest.Pages.Count);
        Assert.Equal(new[] { 2, 3 }, first.Manifest.Pages.Select(page => page.SourcePageNumber).ToArray());
        Assert.Equal(HtmlRenderPageSetMode.Range, first.Manifest.PageSet);
        Assert.Equal(1, first.Manifest.FirstPageIndex);
        Assert.Equal(2, first.Manifest.PageCount);
        Assert.Equal(options.Scale, first.Manifest.RequestedScale);
        Assert.Equal(options.BackgroundColor, first.Manifest.BackgroundColor);
        Assert.Contains(HtmlCapabilityProviderIds.OfficeIMOHtml, first.Manifest.ProviderIds);

        using var archive = new ZipArchive(new MemoryStream(first.Bytes), ZipArchiveMode.Read);
        Assert.Equal(new[] { "pages/page-0001.svg", "pages/page-0002.svg", "manifest.json" },
            archive.Entries.Select(entry => entry.FullName).ToArray());
        ZipArchiveEntry manifestEntry = archive.GetEntry("manifest.json")!;
        using var manifestStream = manifestEntry.Open();
        using JsonDocument json = JsonDocument.Parse(manifestStream);
        Assert.Equal(HtmlRenderArchiveManifest.SchemaId, json.RootElement.GetProperty("schemaId").GetString());
        Assert.Equal("Range", json.RootElement.GetProperty("request").GetProperty("pageSet").GetString());
        Assert.Equal(1, json.RootElement.GetProperty("request").GetProperty("firstPageIndex").GetInt32());
        Assert.Equal(2, json.RootElement.GetProperty("request").GetProperty("pageCount").GetInt32());
        Assert.Equal("#0a141ec8", json.RootElement.GetProperty("request").GetProperty("backgroundColor").GetString());
        Assert.Equal(2, json.RootElement.GetProperty("pages").GetArrayLength());
        Assert.Equal(first.Manifest.Pages[0].Sha256,
            json.RootElement.GetProperty("pages")[0].GetProperty("sha256").GetString());
    }

    [Fact]
    public void PngArchiveUsesTheSharedEncoderAndEnforcesFinalArchiveLimit() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D)
        };
        HtmlRenderResult result = HtmlRenderEngine.Execute(
            HtmlConversionDocument.Parse("<div style='width:60px;height:30px;background:#369'>PNG</div>"),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.Png, options));
        HtmlRenderArchiveResult archive = result.ExportArchive();

        Assert.Equal("image/png", Assert.Single(archive.Manifest.Pages).MimeType);
        Assert.Throws<HtmlRenderArchiveLimitException>(() => result.ExportArchive(
            new HtmlRenderArchiveOptions { MaximumArchiveBytes = 32 }));
        Assert.Throws<InvalidOperationException>(() => HtmlRenderEngine.Execute(
                HtmlConversionDocument.Parse("<p>display list</p>"),
                HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.DisplayList))
            .ExportArchive());
    }

    [Fact]
    public void PngArchiveReportsEncoderScaleLossAndDiagnosticProvenance() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D),
            Scale = 4D,
            MaximumRasterPixels = 1_000,
            RasterOverflowBehavior = OfficeRasterOverflowBehavior.ReduceScale
        };
        HtmlRenderResult result = HtmlRenderEngine.Execute(
                HtmlConversionDocument.Parse("<div style='width:60px;height:30px;background:#369'>PNG</div>"),
                HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.Png, options))
            .WithAdditionalDiagnostics(new[] {
                new HtmlDiagnostic(
                    "ArchiveBoundary",
                    "ARCHIVE_SOURCE",
                    "Container evidence was retained.",
                    HtmlDiagnosticSeverity.Info,
                    "archive.mhtml",
                    "boundary detail",
                    OfficeConversionLossKind.None,
                    sourceLocation: null,
                    targetAddress: "pages/page-0001.png")
            });

        HtmlRenderArchiveResult archive = result.ExportArchive();
        HtmlRenderArchivePage page = Assert.Single(archive.Manifest.Pages);

        Assert.True(page.HasLoss);
        Assert.True(archive.Manifest.HasLoss);
        Assert.Contains(page.EncodingDiagnostics,
            diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.RasterScaleReduced
                          && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        using JsonDocument json = JsonDocument.Parse(archive.Manifest.ToJson());
        JsonElement jsonPage = json.RootElement.GetProperty("pages")[0];
        Assert.True(jsonPage.GetProperty("hasLoss").GetBoolean());
        Assert.Contains(jsonPage.GetProperty("encodingDiagnostics").EnumerateArray(),
            diagnostic => diagnostic.GetProperty("code").GetString() == OfficeImageExportDiagnosticCodes.RasterScaleReduced);
        JsonElement provenance = Assert.Single(json.RootElement.GetProperty("diagnostics").EnumerateArray())
            .GetProperty("provenance");
        Assert.Equal("archive.mhtml", provenance.GetProperty("sourceAddress").GetString());
        Assert.Equal(0, provenance.GetProperty("sourceLine").GetInt32());
        Assert.Equal(0, provenance.GetProperty("sourceColumn").GetInt32());
        Assert.Equal("pages/page-0001.png", provenance.GetProperty("targetAddress").GetString());
    }

}
