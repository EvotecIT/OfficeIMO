using OfficeIMO.Reader;
using OfficeIMO.Reader.DjVu;

namespace OfficeIMO.DjVu.Tests;

public sealed class ReaderAdapterTests {
    internal static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Authored", name);

    [Fact]
    public void StoredUnicodeAndRotationRetainSourceGeometryAndStatus() {
        var document = DjVuDocument.Load(Fixture("reader-book.djvu"));
        Assert.Equal(new[] { DjVuTextStatus.Absent, DjVuTextStatus.Present, DjVuTextStatus.Empty, DjVuTextStatus.Corrupt }, document.Pages.Select(p => p.GetText().Status));
        var text = document.Pages[1].GetText();
        Assert.Contains("Zażółć 😀", text.Text);
        var word = text.Zones[0].Children[0].Children[1];
        Assert.Equal("😀 ", text.Text.Substring(word.CharacterOffset, word.CharacterLength));
        Assert.Equal(5, word.ByteLength);
        Assert.Equal(3, word.CharacterLength);
        var rich = document.ToReadResult();
        Assert.Equal(4, rich.Pages.Count);
        Assert.Equal(new[] { 1, 3 }, rich.OcrCandidates.Select(c => c.Location.Page!.Value));
        Assert.Empty(rich.Assets);
        Assert.All(rich.Blocks, b => Assert.Null(b.Recognition));
        var first = rich.Pages[1].Blocks[0];
        Assert.Equal("Zażółć ", first.Text);
        var bounds = document.Pages[1].GetDisplayBounds(text.Zones[0].Children[0].Children[0].Bounds);
        Assert.Equal(bounds.X * 72.0 / 300, first.Region!.X, 8);
        Assert.Equal(bounds.Y * 72.0 / 300, first.Region.Y, 8);
        Assert.Equal(96 * 72.0 / 300, rich.Pages[1].Width!.Value, 8);
        Assert.Equal(128 * 72.0 / 300, rich.Pages[1].Height!.Value, 8);
        Assert.Contains(rich.Diagnostics, d => d.Code == "djvu.text.corrupt" && d.Location!.Page == 4);
        Assert.Equal(2, rich.Links.First(l => l.Text == "Stored text").DestinationPageNumber);
    }

    [Fact]
    public void BuilderSnapsOptionsAndDetectsSignatureFromCallerStream() {
        var options = new ReaderDjVuOptions { PageNumbers = new[] { 2, 1 } };
        var reader = new OfficeDocumentReaderBuilder().AddDjVuHandler(options).Build();
        options.PageNumbers = new[] { 4 };
        byte[] bytes = File.ReadAllBytes(Fixture("reader-book.djvu"));
        using var stream = new MemoryStream(bytes); stream.Position = 7;
        var rich = reader.ReadDocument(stream, "unknown.bin", new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent, ComputeHashes = true, MaxChars = 3 });
        Assert.True(stream.CanRead);
        Assert.Equal(7, stream.Position);
        Assert.Equal(new[] { 2, 1 }, rich.Pages.Select(p => p.Number!.Value));
        Assert.Equal(ReaderInputKind.DjVu, rich.Kind);
        Assert.Equal(DjVuDocument.Load(bytes).SourceSha256, rich.Source.SourceHash);
        Assert.Equal(bytes.LongLength, rich.Source.LengthBytes);
        Assert.Equal(DjVuDocument.Load(bytes).Pages[1].GetText().Text, string.Concat(rich.Chunks.Select(c => c.Text)));
        Assert.All(rich.Chunks, c => Assert.False(c.Text.Length > 0 && (char.IsLowSurrogate(c.Text[0]) || char.IsHighSurrogate(c.Text[c.Text.Length - 1]))));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SelectedPageOrderSurvivesRichTraversalAndTransport(bool wordZones) {
        byte[] index = File.ReadAllBytes(Fixture(Path.Combine("Shared", "Indirect", "index.djvu")));
        var components = SharedComponentTests.Components();
        byte[] nativeText = SharedComponentTests.Payload(File.ReadAllBytes(Fixture("unicode.djvu")), "TXTz");
        for (int number = 1; number <= 2; number++) {
            string id = "shared-" + number + ".djvu";
            byte[] text = Encoding.UTF8.GetBytes("Page " + number + " stored text");
            var payload = new byte[text.Length + 3];
            payload[0] = (byte)(text.Length >> 16); payload[1] = (byte)(text.Length >> 8); payload[2] = (byte)text.Length;
            text.CopyTo(payload, 3);
            components[id] = SharedComponentTests.AppendChunk(components[id], wordZones ? "TXTz" : "TXTa", wordZones ? nativeText : payload);
        }
        var document = DjVuDocument.Load(index, new DjVuReadOptions { ComponentResolver = (id, _) => components[id] });
        var rich = document.ToReadResult(new ReaderDjVuOptions { PageNumbers = new[] { 2, 1 } }, readerOptions: new ReaderOptions { MaxChars = 3 });
        AssertOrder(rich);
        var restored = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(rich));
        AssertOrder(restored);
        restored.Blocks = Array.Empty<OfficeDocumentBlock>();
        foreach (var page in restored.Pages) page.Blocks = Array.Empty<OfficeDocumentBlock>();
        Assert.Equal(new[] { 2, 1 }, restored.EnumerateContent().Select(item => item.Location!.Page!.Value).Distinct());

        static void AssertOrder(OfficeDocumentReadResult result) {
            Assert.Equal(new[] { 2, 1 }, result.Pages.Select(page => page.Number!.Value));
            Assert.Equal(new[] { 2, 1 }, result.Chunks.Select(chunk => chunk.Location.Page!.Value).Distinct());
            Assert.Equal(new[] { 2, 1 }, result.EnumerateBlocks().Select(block => block.Location!.Page!.Value).Distinct());
            Assert.Equal(new[] { 2, 1 }, result.EnumerateContent().Select(item => item.Location!.Page!.Value).Distinct());
            Assert.Equal(result.Blocks.Select(block => block.Text), result.EnumerateBlocks().Select(block => block.Text));
            Assert.Equal(new long?[] { 1, 2 }, result.Pages.Select(page => page.Location.LogicalOrder));
            foreach (var page in result.Pages) {
                Assert.All(page.Blocks, block => Assert.Equal(page.Location.LogicalOrder, block.Location!.LogicalOrder));
                Assert.All(result.Chunks.Where(chunk => chunk.Location.Page == page.Number), chunk => Assert.Equal(page.Location.LogicalOrder, chunk.Location.LogicalOrder));
            }
        }
    }

    [Fact]
    public void DjVuTransportRequiresItsVersionedSchemaAndRetainsGeometry() {
        var rich = DjVuDocument.Load(Fixture("reader-book.djvu")).ToReadResult();
        string json = OfficeDocumentReadResultJson.Serialize(rich);
        var restored = OfficeDocumentReadResultJson.Deserialize(json);
        Assert.Equal(11, restored.SchemaVersion);
        Assert.Equal(ReaderInputKind.DjVu, restored.Kind);
        Assert.Equal(rich.Blocks[0].Region!.X, restored.Blocks[0].Region!.X);
        Assert.Equal(4, restored.GetTotalPageCount());
        Assert.Equal(OfficeDocumentPageProvenance.Native, restored.GetPageProvenance());
        rich.SchemaVersion = 10;
        Assert.Throws<System.Text.Json.JsonException>(() => OfficeDocumentReadResultJson.Serialize(rich));
    }

    [Fact]
    public void ImagesAreOptInAndAggregateLimitsApplyBeforeProjection() {
        var document = DjVuDocument.Load(Fixture("reader-book.djvu"));
        var rich = document.ToReadResult(new ReaderDjVuOptions { ImageMode = ReaderDjVuImageMode.MissingTextPages });
        Assert.Equal(new[] { 1, 3 }, rich.Assets.Select(a => a.Location.Page!.Value));
        Assert.All(rich.OcrCandidates, c => Assert.NotNull(c.AssetId));
        Assert.All(rich.Assets, a => { Assert.Equal("image/png", a.MediaType); Assert.True(a.PayloadBytes!.Length > 8); });
        var readerOptions = ReaderOptions.CreateSafeIngestion(); readerOptions.MaxChars = 3; readerOptions.ResourceLimits!.MaxChunks = 1;
        Assert.Throws<ReaderResourceLimitException>(() => document.ToReadResult(readerOptions: readerOptions));
        Assert.Throws<ReaderResourceLimitException>(() => document.ToReadResult(new ReaderDjVuOptions { ImageMode = ReaderDjVuImageMode.AllPages, MaxPageImages = 1 }));
        Assert.Throws<OperationCanceledException>(() => document.ToReadResult(cancellationToken: new CancellationToken(true)));
    }

    [Fact]
    public void CompressedNavigationPreservesHierarchyAndResolvesOnlyLocalTargets() {
        var document = DjVuDocument.Load(Fixture("reader-book.djvu"));
        Assert.Null(document.BookmarkDiagnostic);
        Assert.Equal(3, document.Bookmarks.Count);
        Assert.Equal("Group 😀", document.Bookmarks[0].Title);
        Assert.Null(document.Bookmarks[0].PageNumber);
        Assert.Equal(new int?[] { 2, 1 }, document.Bookmarks[0].Children.Select(c => c.PageNumber));
        Assert.Equal("https://example.org/archive", document.Bookmarks[1].Target);
        Assert.Null(document.Bookmarks[1].PageNumber);
        Assert.Equal(3, document.Bookmarks[2].PageNumber);
        Assert.Throws<DjVuResourceLimitException>(() => DjVuDocument.Load(Fixture("reader-book.djvu"), new DjVuReadOptions { MaxBookmarks = 4 }));
        Assert.Throws<DjVuResourceLimitException>(() => DjVuDocument.Load(Fixture("reader-book.djvu"), new DjVuReadOptions { MaxBookmarkDepth = 1 }));
    }

    [Theory]
    [InlineData("count", nameof(ReaderDjVuOptions.MaxPageImages))]
    [InlineData("page", nameof(ReaderDjVuOptions.MaxPageImageBytes))]
    [InlineData("total", nameof(ReaderDjVuOptions.MaxTotalPageImageBytes))]
    [InlineData("reader", nameof(ReaderResourceLimits.MaxAssetBytes))]
    public void ImageBudgetFailuresIdentifyTheActualLimit(string budget, string expected) {
        var document = DjVuDocument.Load(Fixture("reader-book.djvu"));
        var images = new ReaderDjVuOptions { ImageMode = ReaderDjVuImageMode.AllPages };
        var reader = new ReaderOptions();
        switch (budget) {
            case "count": images.MaxPageImages = 0; break;
            case "page": images.MaxPageImageBytes = 8; break;
            case "total": images.MaxTotalPageImageBytes = 8; break;
            case "reader": reader.ResourceLimits = new ReaderResourceLimits { MaxAssetBytes = 0 }; break;
        }
        var error = Assert.Throws<ReaderResourceLimitException>(() => document.ToReadResult(images, readerOptions: reader));
        Assert.Equal(expected, error.LimitName);
        Assert.Equal(budget == "count" || budget == "reader" ? 0 : 8, error.Maximum);
        Assert.Empty(document.ToReadResult(new ReaderDjVuOptions { MaxPageImages = 0 }).Assets);
    }
}
