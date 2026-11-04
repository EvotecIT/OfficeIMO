using OfficeIMO.Drawing;
using OfficeIMO.Publisher;
using OfficeIMO.Reader.Publisher;
using OfficeIMO.Tests;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPublisherFrameCreationTests {
    [Fact]
    public void InMemoryPublisherProjectionKeepsOneCompleteStoryAcrossImportedAndAuthoredFrames() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Publisher.Tests", "Fixtures",
            "NativeComparison", "TextBlank", "native-textblank.pub");
        byte[] input = File.ReadAllBytes(path);
        PublisherDocument document = PublisherDocument.Load(input);
        PublisherTextFrame original = Assert.Single(document.Pages[0].TextFrames);
        PublisherTextStory originalStory = Assert.Single(document.TextStories);
        var sourceReport = document.SourceReadReport;
        const string heading = "Continued publication";
        string body = string.Concat(Enumerable.Repeat("Continuation text stays complete. ", 60));
        document.ReplaceStoryParagraphs(original.StoryId, new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(heading, 16, OfficeColor.Black, bold: true, fontFamily: "Arial") }),
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(body, 11, OfficeColor.Black, fontFamily: "Arial") })
        });
        document.SetTextFrameBounds(original.Id, 54, 108, 228, 120);
        PublisherTextFrame beforeAppend = Assert.Single(document.Pages[0].TextFrames);
        Assert.True(beforeAppend.HasOverflow);
        PublisherTextFrame appended = document.AddLinkedTextFrame(original.Id, document.Pages[1].Id, 54, 72, 468, 600);
        PublisherTextFrame first = Assert.Single(document.Pages[0].TextFrames);
        PublisherTextFrame last = Assert.Single(document.Pages[1].TextFrames);
        Assert.Equal(appended.Id, last.Id);
        Assert.NotEqual(original.Id, last.Id);
        Assert.Equal(original.StoryId, last.StoryId);
        Assert.Equal(last.Id, first.NextFrameId);
        Assert.Equal(first.Id, last.PreviousFrameId);
        Assert.Equal(first.TextLength, last.TextStart);
        Assert.True(last.TextLength > 0);
        Assert.False(last.HasOverflow);
        Assert.Null(beforeAppend.NextFrameId);
        Assert.True(beforeAppend.HasOverflow);
        Assert.Equal("Publisher text\n", originalStory.Text);
        Assert.Same(sourceReport, document.SourceReadReport);

        string expected = heading + "\n" + body + "\n";
        OfficeDocumentReadResult result = document.ToOfficeDocumentReadResult("authored-continuation.pub",
            new ReaderOptions { ComputeHashes = true, MaxChars = 256 });
        Assert.Equal(expected, string.Concat(result.Blocks.Select(block => block.Text)));
        Assert.Equal(expected, string.Concat(result.Chunks.Select(chunk => chunk.Text)));
        Assert.Equal(2, result.Pages.Count);
        Assert.Equal(OfficeDocumentPageProvenance.Native, result.GetPageProvenance());
        Assert.Null(result.Source.SourceHash);
        Assert.Null(result.Source.LengthBytes);
        Assert.All(result.Blocks, block => Assert.Null(block.Location.Page));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "PUB_READER_LAYOUT_OMITTED");

        document.ReplaceStoryParagraphs(original.StoryId, new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Short", 11, OfficeColor.Black, fontFamily: "Arial") })
        });
        OfficeDocumentReadResult shorter = document.ToOfficeDocumentReadResult("authored-continuation.pub");
        Assert.Equal("Short\n", string.Concat(shorter.Blocks.Select(block => block.Text)));
        Assert.Equal(0, Assert.Single(document.Pages[1].TextFrames).TextLength);
        Assert.Equal(expected, string.Concat(result.Blocks.Select(block => block.Text)));
        Assert.Equal(input, File.ReadAllBytes(path));
    }
}
