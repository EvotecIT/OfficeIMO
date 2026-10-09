namespace OfficeIMO.DjVu.Tests;

public sealed class RenderWorkLimitsTests {
    [Fact]
    public void ProgressiveSlicesAreBoundedAcrossChunksEvenWithImplicitEntropy() {
        byte[] source = SharedComponentTests.WithoutChunks(File.ReadAllBytes(ReaderAdapterTests.Fixture("palette.djvu")), "Sjbz", "FGbz", "BG44");
        source = SharedComponentTests.AppendChunk(source, "BG44", new byte[] { 0, 255, 129, 2, 0, 128, 0, 96, 0 });
        source = SharedComponentTests.AppendChunk(source, "BG44", new byte[] { 1, 2 });
        var page = DjVuDocument.Load(source).Pages[0];
        Assert.Equal(nameof(DjVuReadOptions.MaxIw44Slices), Assert.Throws<DjVuResourceLimitException>(() => page.Render()).LimitName);
    }

    [Fact]
    public void CoefficientWorkAccumulatesAcrossProgressiveChunks() {
        byte[] source = SharedComponentTests.WithoutChunks(File.ReadAllBytes(ReaderAdapterTests.Fixture("palette.djvu")), "Sjbz", "FGbz", "BG44");
        source = SharedComponentTests.AppendChunk(source, "BG44", new byte[] { 0, 1, 129, 2, 0, 128, 0, 96, 0 });
        source = SharedComponentTests.AppendChunk(source, "BG44", new byte[] { 1, 2 });
        var page = DjVuDocument.Load(source, new DjVuReadOptions { MaxIw44CoefficientSamples = 400 }).Pages[0];
        Assert.Equal(nameof(DjVuReadOptions.MaxIw44CoefficientSamples), Assert.Throws<DjVuResourceLimitException>(() => page.Render()).LimitName);
    }

    [Fact]
    public void MaskPaintBudgetCountsAllPlacementsRatherThanOnlyLargestBitmap() {
        var page = DjVuDocument.Load(ReaderAdapterTests.Fixture("palette.djvu")).Pages[0];
        var mask = page.DecodeMask(new DjVuReadBudget(new DjVuReadOptions(), default));
        long largest = mask.Placements.Max(p => (long)p.Bitmap.Width * p.Bitmap.Height);
        Assert.True(mask.Placements.Sum(p => (long)p.Bitmap.Width * p.Bitmap.Height) > largest);
        Assert.Equal(nameof(DjVuRenderOptions.MaxMaskPaintSamples), Assert.Throws<DjVuResourceLimitException>(() =>
            page.Render(new DjVuRenderOptions { MaxMaskPaintSamples = largest })).LimitName);
        Assert.Equal(PageRenderTests.ReadPpm(ReaderAdapterTests.Fixture("palette-reference.ppm")).GetPixels(), page.Render().Image.GetPixels());
    }

    [Fact]
    public void Jb2DecodeBudgetIncludesSharedDictionariesBeforePagePainting() {
        var page = DjVuDocument.Load(ReaderAdapterTests.Fixture(Path.Combine("Shared", "shared.djvu")),
            new DjVuReadOptions { MaxJb2DecodedSamples = 1 }).Pages[0];
        Assert.Equal(nameof(DjVuReadOptions.MaxJb2DecodedSamples), Assert.Throws<DjVuResourceLimitException>(() => page.Render()).LimitName);
    }
}
