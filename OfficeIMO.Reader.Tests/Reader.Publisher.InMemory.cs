using OfficeIMO.Drawing;
using OfficeIMO.Publisher;
using OfficeIMO.Reader.Publisher;
using OfficeIMO.Tests;
using System.Globalization;
using System.Threading;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPublisherInMemoryTests {
    [Fact]
    public void InMemoryPublisherProjectionUsesEditedLinkedStoriesAndCurrentPageSnapshots() {
        string path = NewsletterPath();
        PublisherDocument document = PublisherDocument.Load(path);
        PublisherTextFrame[] originalFrames = document.Pages.SelectMany(page => page.TextFrames)
            .Where(frame => frame.PreviousFrameId.HasValue || frame.NextFrameId.HasValue).OrderBy(frame => frame.Order).ToArray();
        Assert.Equal(2, originalFrames.Length);
        Assert.NotEqual(originalFrames[0].PageId, originalFrames[1].PageId);
        uint storyId = originalFrames[0].StoryId;
        PublisherTextStory originalStory = Assert.Single(document.TextStories, story => story.Id == storyId);
        OfficeDocumentReadResult original = document.ToOfficeDocumentReadResult(path,
            new ReaderOptions { MaxChars = 256, ComputeHashes = true });
        OfficeDocumentReadResult fileResult = new OfficeDocumentReaderBuilder().AddPublisherHandler().Build()
            .ReadDocument(path, new ReaderOptions { ComputeHashes = true });
        Assert.NotNull(fileResult.Source.SourceHash);

        const string heading = "Edited **Publisher** story";
        string body = string.Join(" ", Enumerable.Repeat("Current semantic content", 40)) + " complete ending.";
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("•", 12, OfficeColor.Black, fontFamily: "Arial"), 0);
        var paragraphs = new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(heading, 16, OfficeColor.Black, bold: true, fontFamily: "Arial") }, OfficeTextAlignment.Center),
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(body, 12, OfficeColor.Black, italic: true, fontFamily: "Arial") }, label)
        };
        document.ReplaceStoryParagraphs(storyId, paragraphs);
        for (int index = 0; index < originalFrames.Length; index++)
            document.SetTextFrameBounds(originalFrames[index].Id, 72 + index * 6, 90 + index * 12, 420, 36);

        PublisherTextStory currentStory = Assert.Single(document.TextStories, story => story.Id == storyId);
        string expectedText = heading + "\n" + body + (originalStory.Text.EndsWith("\n", StringComparison.Ordinal) ? "\n" : string.Empty);
        Assert.Equal(expectedText, currentStory.Text);
        Assert.True(currentStory.Paragraphs[0].Runs[0].Bold);
        Assert.True(currentStory.Paragraphs[1].Runs[0].Italic);
        PublisherTextFrame[] currentFrames = document.Pages.SelectMany(page => page.TextFrames)
            .Where(frame => frame.StoryId == storyId).OrderBy(frame => frame.Order).ToArray();
        Assert.Equal(originalFrames.Select(frame => frame.Id), currentFrames.Select(frame => frame.Id));
        Assert.Equal(originalFrames.Select(frame => frame.PageId), currentFrames.Select(frame => frame.PageId));
        Assert.Equal(currentFrames[0].Id, currentFrames[1].PreviousFrameId);
        Assert.Equal(currentFrames[1].Id, currentFrames[0].NextFrameId);
        for (int index = 0; index < currentFrames.Length; index++) {
            Assert.Equal(72 + index * 6, currentFrames[index].X);
            Assert.Equal(90 + index * 12, currentFrames[index].Y);
            Assert.Equal(420, currentFrames[index].Width);
            Assert.Equal(36, currentFrames[index].Height);
        }

        OfficeDocumentReadResult result = document.ToOfficeDocumentReadResult(path,
            new ReaderOptions { MaxChars = 256, ComputeHashes = true });
        string anchor = "publisher-story-" + storyId.ToString(CultureInfo.InvariantCulture);
        OfficeDocumentBlock[] blocks = result.Blocks.Where(block => block.Location.BlockAnchor == anchor).ToArray();
        ReaderChunk[] chunks = StoryChunks(result, anchor);
        Assert.Equal(ReaderInputKind.Publisher, result.Kind);
        Assert.Equal(expectedText, string.Concat(blocks.Select(block => block.Text)));
        Assert.Equal(expectedText, string.Concat(chunks.Select(chunk => chunk.Text)));
        Assert.Equal("paragraph", blocks[0].Kind);
        Assert.Equal("list-item", blocks[1].Kind);
        Assert.Equal("•", blocks[1].Marker);
        Assert.Contains("- Current semantic content", result.Markdown);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "PUB_READER_LAYOUT_OMITTED");
        Assert.Equal(document.Pages.Count, result.Pages.Count);
        Assert.Equal(OfficeDocumentPageProvenance.Native, result.GetPageProvenance());
        for (int index = 0; index < document.Pages.Count; index++) {
            Assert.Equal(index + 1, result.Pages[index].Number);
            Assert.Equal(index + 1, result.Pages[index].Location.Page);
            Assert.Equal(document.Pages[index].Width, result.Pages[index].Width);
            Assert.Equal(document.Pages[index].Height, result.Pages[index].Height);
            Assert.Equal(path, result.Pages[index].Location.Path);
        }
        // A complete linked story spans physical frames; semantic paragraphs
        // retain logical anchors instead of claiming one frame's page citation.
        Assert.All(blocks, block => { Assert.Equal(path, block.Location.Path); Assert.Null(block.Location.Page); });
        Assert.All(chunks, chunk => { Assert.Equal(path, chunk.Location.Path); Assert.Null(chunk.Location.Page); });
        Assert.Equal(original.Source.SourceId, result.Source.SourceId);
        Assert.Null(result.Source.SourceHash);
        Assert.Null(result.Source.LengthBytes);
        Assert.Null(result.Source.LastWriteUtc);
        Assert.All(result.Chunks, chunk => {
            Assert.Equal(result.Source.SourceId, chunk.SourceId);
            Assert.Null(chunk.SourceHash);
            Assert.Null(chunk.SourceLengthBytes);
            Assert.Null(chunk.SourceLastWriteUtc);
            Assert.NotNull(chunk.ChunkHash);
        });
        Assert.NotEqual(StoryChunks(original, anchor)[0].ChunkHash, chunks[0].ChunkHash);
        Assert.Equal(originalStory.Text, string.Concat(StoryChunks(original, anchor).Select(chunk => chunk.Text)));
    }

    [Theory]
    [InlineData(null, "document.pub")]
    [InlineData("  edited.pub  ", "edited.pub")]
    public void InMemoryPublisherProjectionNormalizesLogicalNamesAndReaderOptions(string? sourceName, string expectedName) {
        PublisherDocument document = PublisherDocument.Load(NewsletterPath());
        OfficeDocumentReadResult result = document.ToOfficeDocumentReadResult(sourceName,
            new ReaderOptions { MaxChars = 0, ComputeHashes = false });
        Assert.Equal(expectedName, result.Source.Path);
        Assert.NotEmpty(result.Chunks);
        Assert.All(result.Chunks, chunk => {
            Assert.Equal(expectedName, chunk.Location.Path);
            Assert.True(chunk.Text.Length <= 256);
            Assert.True(chunk.Markdown!.Length <= 256);
            Assert.Null(chunk.ChunkHash);
            Assert.Null(chunk.SourceHash);
        });
    }

    [Fact]
    public void InMemoryPublisherProjectionEnforcesSharedOutputBudgets() {
        PublisherDocument document = PublisherDocument.Load(NewsletterPath());
        Assert.Throws<ReaderResourceLimitException>(() => document.ToOfficeDocumentReadResult(readerOptions:
            new ReaderOptions { MaxChars = 256, ResourceLimits = new ReaderResourceLimits { MaxChunks = 1 } }));
    }

    [Fact]
    public void InMemoryPublisherProjectionRejectsNullAndHonorsCancellation() {
        Assert.Throws<ArgumentNullException>(() => ((PublisherDocument)null!).ToOfficeDocumentReadResult());
        PublisherDocument document = PublisherDocument.Load(NewsletterPath());
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToOfficeDocumentReadResult(cancellationToken: cancelled.Token));
    }

    [Fact]
    public void InMemoryPublisherProjectionKeepsUnsupportedProjectionModesExplicit() {
        PublisherDocument document = PublisherDocument.Load(NewsletterPath());
        Assert.Throws<NotSupportedException>(() => document.ToOfficeDocumentReadResult(readerOptions:
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata }));
    }

    private static ReaderChunk[] StoryChunks(OfficeDocumentReadResult result, string anchor) => result.Chunks
        .Where(chunk => chunk.Location.BlockAnchor?.StartsWith(anchor + "-p", StringComparison.Ordinal) == true).ToArray();

    private static string NewsletterPath() => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Publisher.Tests",
        "Fixtures", "NativeComparison", "Newsletter", "native-newsletter.pub");
}
