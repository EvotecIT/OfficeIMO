using OfficeIMO.Publisher;
using OfficeIMO.Markdown;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Publisher;
using System.Security.Cryptography;
using System.Threading;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPublisherTests {
    [Theory]
    [InlineData("Sample.pub")]
    [InlineData("SampleNewsletter.pub")]
    [InlineData("SampleBrochure.pub")]
    public void CompleteNativeStoriesAreProjectedOnceWithPhysicalInventoryAndLosses(string fixture) {
        byte[] bytes = Fixture(fixture);
        PublisherDocument native = PublisherDocument.Load(bytes);
        OfficeDocumentReadResult result = CreateReader().ReadDocument(bytes, "source.pub", new ReaderOptions { MaxChars = 256 });
        Assert.Equal(ReaderInputKind.Publisher, result.Kind);
        Assert.Equal(native.Pages.Count, result.Pages.Count); Assert.Equal(OfficeDocumentPageProvenance.Native, result.GetPageProvenance());
        foreach (PublisherTextStory story in native.TextStories) {
            string anchor = "publisher-story-" + story.Id;
            OfficeDocumentBlock[] blocks = result.Blocks.Where(block => block.Location.BlockAnchor == anchor).ToArray();
            Assert.Equal(story.Text, string.Concat(blocks.Select(block => block.Text)));
            Assert.All(blocks, block => Assert.Null(block.Location.Page));
        }
        Assert.Equal(string.Concat(result.Blocks.Select(block => block.Text)), string.Concat(result.Chunks.Select(chunk => chunk.Text)));
        Assert.Equal(result.Blocks.Select(block => block.Id), result.EnumerateBlocks().Select(block => block.Id));
        Assert.Equal(result.Blocks.Select(block => block.Id), result.EnumerateContent().Select(item => item.Block!.Id));
        OfficeDocumentReadResult transported = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(result));
        Assert.Equal(result.Blocks.Select(block => block.Id), transported.EnumerateBlocks().Select(block => block.Id));
        Assert.All(native.ReadReport.FidelityDiagnostics, finding => Assert.Contains(result.Diagnostics,
            retained => retained.Code == finding.Code && retained.Attributes["lossKind"] == finding.LossKind.ToString()));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "PUB_READER_LAYOUT_OMITTED");
        Assert.Equal(native.Images.Count, result.Assets.Count); Assert.All(result.Assets, asset => Assert.Null(asset.PayloadBytes));
        if (fixture == "SampleNewsletter.pub") {
            Assert.Contains(result.Blocks, block => block.Kind == "list-item" && block.Marker == "•");
            Assert.Contains("- Jay Adams", result.Markdown);
        }
        Assert.All(result.Chunks, chunk => { Assert.True(chunk.Text.Length <= 126); Assert.True(chunk.Markdown!.Length <= 256); });
    }

    [Theory]
    [InlineData("source.pub")]
    [InlineData("renamed.bin")]
    public void MetadataDetectionFindsPublisherWithoutParsingItsDrawing(string name) {
        byte[] bytes = Fixture("Sample.pub");
        OfficeDocumentReader reader = CreateReader();
        ReaderDetectionResult detected = reader.Detect(bytes, name, new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent });
        Assert.Equal(ReaderInputKind.Publisher, detected.ContentKind);
        Assert.Contains("container:ole-publisher-quill", detected.Evidence);
        Assert.Equal("application/x-mspublisher", detected.MediaType);
        OfficeDocumentReadResult result = reader.ReadDocument(bytes, name);
        Assert.Equal(ReaderInputKind.Publisher, result.Kind);
        Assert.Equal(ReaderInputKind.Publisher,
            OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(result)).Kind);
        ReaderHandlerCapability capability = Assert.Single(reader.GetCapabilities(), item => item.Id == "officeimo.reader.publisher");
        Assert.Equal(ReaderFormatSupport.ReadConvert, Assert.Single(capability.FormatQualifications).Support);
    }

    [Fact]
    public void CallerStreamPositionAndCapturedSourceIdentityArePreserved() {
        byte[] bytes = Fixture("Sample.pub");
        using var stream = new MemoryStream(bytes); stream.Position = 7;
        var result = CreateReader().ReadDocument(stream, "source.pub", new ReaderOptions { ComputeHashes = true });
        Assert.Equal(7, stream.Position); Assert.True(stream.CanRead);
        Assert.Equal(bytes.LongLength, result.Source.LengthBytes);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(), result.Source.SourceHash);
        Assert.All(result.Chunks, chunk => { Assert.Equal(result.Source.SourceHash, chunk.SourceHash); Assert.NotNull(chunk.ChunkHash); });
    }

    [Fact]
    public void PresetAndDetachedOptionsUseTheNativeOwnerAndOptionalImageCopies() {
        var options = new ReaderPublisherOptions { ReadOptions = new() { MaximumPages = 8 }, IncludeImagePayloads = true };
        OfficeDocumentReader reader = CreateReader(options);
        options.ReadOptions.MaximumPages = 1; options.IncludeImagePayloads = false;
        byte[] input = Fixture("SampleNewsletter.pub");
        var result = reader.ReadDocument(input, "source.pub");
        PublisherDocument native = PublisherDocument.Load(input);
        Assert.NotEmpty(result.Assets);
        for (int index = 0; index < native.Images.Count; index++) Assert.Equal(native.Images[index].GetBytes(), result.Assets[index].PayloadBytes);
        var preset = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build().ReadDocument(input, "source.pub");
        Assert.Equal(ReaderInputKind.Publisher, preset.Kind); Assert.NotEmpty(preset.Blocks);
    }

    [Theory]
    [InlineData("Sample98.pub")]
    [InlineData("Sample2000.pub")]
    public void UnsupportedNativeProfilesAreNotRoutedToTextSalvage(string fixture) =>
        Assert.Throws<NotSupportedException>(() => CreateReader().ReadDocument(Fixture(fixture), "source.pub"));

    [Fact]
    public void ReaderCeilingsAndCancellationReachTheSharedAndNativeBoundaries() {
        byte[] bytes = Fixture("Sample.pub");
        Assert.Throws<IOException>(() => CreateReader().ReadDocument(bytes, "source.pub", new ReaderOptions { MaxInputBytes = 32 }));
        Assert.Throws<InvalidDataException>(() => CreateReader(new() { ReadOptions = new() { MaximumPages = 1 } }).ReadDocument(bytes, "source.pub"));
        Assert.Throws<ReaderResourceLimitException>(() => CreateReader().ReadDocument(bytes, "source.pub",
            new ReaderOptions { MaxChars = 10, ResourceLimits = new() { MaxChunks = 1 } }));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => CreateReader().ReadDocument(bytes, "source.pub", cancellationToken: cancelled.Token));
    }

    [Fact]
    public void LiteralSourcePunctuationDoesNotActivateMarkdownOrHtml() {
        const string text = "[click](https://example.test) <script>alert(1)</script> # title *value*";
        string escaped = ReaderMarkdownEscaping.EscapeLiteral(text);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(MarkdownReader.Parse(escaped).Blocks));
        var literal = new System.Text.StringBuilder();
        ((IPlainTextMarkdownInline)paragraph.Inlines).AppendPlainText(literal);
        Assert.Equal(text, literal.ToString());
        string html = ((IMarkdownBlock)paragraph).RenderHtml();
        Assert.Contains("&lt;script&gt;", html);
        Assert.DoesNotContain("<script>", html); Assert.DoesNotContain("href=", html);
        Assert.DoesNotContain("<strong>", html); Assert.DoesNotContain("<h", html);
    }

    private static OfficeDocumentReader CreateReader(ReaderPublisherOptions? options = null) =>
        new OfficeDocumentReaderBuilder().AddPublisherHandler(options).Build();
    private static byte[] Fixture(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "PublisherFixtures", name));
}
