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

    [Fact]
    public void Jb2CommentWorkAccumulatesAcrossRecordsBeforeDecodingTheirBytes() {
        string source = ReaderAdapterTests.Fixture("comments-page.djvu");
        var limited = DjVuDocument.Load(source, new DjVuReadOptions { MaxJb2CommentBytes = 8191 });
        Assert.Equal(nameof(DjVuReadOptions.MaxJb2CommentBytes), Assert.Throws<DjVuResourceLimitException>(() => limited.Pages[0].Render()).LimitName);
        var accepted = DjVuDocument.Load(source, new DjVuReadOptions { MaxJb2CommentBytes = 8192 });
        Assert.DoesNotContain(accepted.Pages[0].Render().Image.GetPixels(), value => value != 255);
    }

    [Fact]
    public void Jb2CommentsInInheritedDictionaryAndPageShareOneWorkBudget() {
        byte[] dictionary = File.ReadAllBytes(ReaderAdapterTests.Fixture("comments-dictionary.jb2"));
        var budget = new DjVuReadBudget(new DjVuReadOptions { MaxJb2CommentBytes = 8192 }, default);
        var chunk = new DjVuChunk(dictionary, "Djbz", null, 0, dictionary.Length, new System.Collections.Generic.List<DjVuChunk>());
        var inherited = new Jb2Decoder(chunk, budget).Decode(true).Library;
        var page = DjVuDocument.Load(ReaderAdapterTests.Fixture("comments-page.djvu")).Pages[0];
        var image = page.Document.PageChunks(page.Component, default).Single(c => c.Id == "Sjbz");
        Assert.Equal(nameof(DjVuReadOptions.MaxJb2CommentBytes), Assert.Throws<DjVuResourceLimitException>(() =>
            new Jb2Decoder(image, budget, inherited).Decode()).LimitName);
    }

    [Fact]
    public void EmptyJb2SymbolsPreservePlacementGeometryAndTrimToEmptyLibraryEntries() {
        var page = DjVuDocument.Load(ReaderAdapterTests.Fixture("zero-area.djvu"), new DjVuReadOptions { MaxJb2DecodedSamples = 1 }).Pages[0];
        var mask = page.DecodeMask(new DjVuReadBudget(page.Document.ReadOptions, default));
        Assert.Equal(1000, mask.Library.Count);
        Assert.All(mask.Library, bitmap => {
            Assert.Equal((0, 0), (bitmap.Width, bitmap.Height));
            Assert.Empty(bitmap.Pixels);
        });
        var placement = Assert.Single(mask.Placements);
        Assert.Equal((0, -65534, 0, 65535), (placement.X, placement.Y, placement.Bitmap.Width, placement.Bitmap.Height));
        Assert.DoesNotContain(page.Render().Image.GetPixels(), value => value != 255);
    }

    [Fact]
    public void FullyClippedReusedMasksDoNotConsumeVisiblePaintSamples() {
        var page = DjVuDocument.Load(ReaderAdapterTests.Fixture("clipped-masks.djvu")).Pages[0];
        var image = page.Render(new DjVuRenderOptions { Region = new DjVuRectangle(1, 0, 1, 65535), MaxMaskPaintSamples = 1 }).Image;
        Assert.Equal((1, 65535), (image.Width, image.Height));
        Assert.DoesNotContain(image.GetPixels(), value => value != 255);
    }
}
