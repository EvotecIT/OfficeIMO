using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Drawing;

namespace OfficeIMO.Studio.Tests;

public sealed class PageSceneCoordinatorTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public async Task ReusesScenesAndEvictsLeastRecentlyUsedEntries() {
        int loads = 0;
        using var coordinator = new PageSceneCoordinator(
            (page, _) => {
                loads++;
                return Task.FromResult(TestPdfPageScenes.Create(page));
            },
            maximumEntries: 2,
            maximumElements: 10);

        PdfPageScene first = await coordinator.GetPageAsync(1, CancellationToken.None);
        Assert.Same(first, await coordinator.GetPageAsync(1, CancellationToken.None));
        await coordinator.GetPageAsync(2, CancellationToken.None);
        await coordinator.GetPageAsync(3, CancellationToken.None);

        Assert.Equal(3, loads);
        Assert.Equal(2, coordinator.CachedEntryCount);
    }

    [Fact]
    public void AdvancedTextProducesAnExplicitAvaloniaFallbackReason() {
        var drawing = new OfficeDrawing(300, 200);
        drawing.AddText(
            "padded",
            10,
            10,
            120,
            30,
            padding: new OfficeTextPadding(2, 2, 2, 2));

        IReadOnlyList<string> reasons = OfficeDrawingAvaloniaRenderer.AnalyzeRasterFallback(drawing);

        Assert.NotEmpty(reasons);
        Assert.Contains(reasons, reason => reason.Contains("text", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void SceneEstimateIncludesPatternFontAndSoftMaskPayloads() {
        var baseline = TestPdfPageScenes.Create();
        byte[] font = File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory,
            "Fixtures",
            "SourceSansPro-Regular.otf"));
        var drawing = new OfficeDrawing(612, 792)
            .AddFont("Source Sans Pro", font)
            .AddImagePattern(
                TinyPng,
                "image/png",
                new OfficeImagePatternLayout(
                    new OfficeImagePlacement(0, 0, 100, 100),
                    new OfficeImagePlacement(0, 0, 10, 10)));
        var maskDrawing = new OfficeDrawing(100, 100)
            .AddImagePattern(
                TinyPng,
                "image/png",
                new OfficeImagePatternLayout(
                    new OfficeImagePlacement(0, 0, 100, 100),
                    new OfficeImagePlacement(0, 0, 10, 10)));
        drawing.AddEffectDrawing(
            new OfficeDrawing(100, 100),
            OfficeTransform.Identity,
            OfficeBlendMode.Normal,
            new OfficeDrawingSoftMask(maskDrawing));

        PdfPageScene scene = TestPdfPageScenes.Create(drawing: drawing);

        Assert.True(scene.EstimatedBytes >= baseline.EstimatedBytes + font.LongLength + (2L * TinyPng.LongLength));
    }

    [Fact]
    public async Task ByteLimitRejectsSceneWhoseRetainedPayloadIsOversized() {
        byte[] payload = new byte[4096];
        var drawing = new OfficeDrawing(100, 100)
            .AddImagePattern(
                payload,
                "application/octet-stream",
                new OfficeImagePatternLayout(
                    new OfficeImagePlacement(0, 0, 100, 100),
                    new OfficeImagePlacement(0, 0, 100, 100)));
        using var coordinator = new PageSceneCoordinator(
            (page, _) => Task.FromResult(TestPdfPageScenes.Create(page, drawing: drawing)),
            maximumBytes: 2048);

        await coordinator.GetPageAsync(1, CancellationToken.None);

        Assert.Equal(0, coordinator.CachedEntryCount);
        Assert.Equal(0, coordinator.CachedByteCount);
    }
}
