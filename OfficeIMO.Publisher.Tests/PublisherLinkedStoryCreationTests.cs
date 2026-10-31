using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherLinkedStoryCreationTests {
    private const string Sentence = "Linked stories preserve reading order across publication pages. ";

    [Theory]
    [InlineData("LinkedStory/native-linkedstory.pub", false)]
    [InlineData("LinkedStoryOverflow/native-linkedstoryoverflow.pub", true)]
    [InlineData("GeneratedLinkedStory/generated-linkedstory.pub", false)]
    [InlineData("GeneratedLinkedStory/native-normalized.pub", false)]
    public void Three_frame_catalog_story_retains_cross_page_order_and_opaque_layout_through_repeated_text_saves(string fixture, bool overflow) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory story = Assert.Single(document.TextStories, value => value.Text.StartsWith("Linked stories", StringComparison.Ordinal));
        Assert.Equal(Text(overflow ? 24 : 6) + "\n", story.Text);
        PublisherTextFrame[] original = Frames(document, story.Id);
        Assert.Equal(3, original.Length);
        Assert.Equal(new[] { document.Pages[0].Id, document.Pages[2].Id, document.Pages[1].Id }, original.Select(frame => frame.PageId));
        Assert.Equal(new uint[] { 0, 1, 2 }, original.Select(frame => frame.Order));
        Assert.Equal(overflow, original[2].HasOverflow);
        Assert.Equal("Independent story\n", Assert.Single(document.TextStories, value => value.Id != story.Id).Text);

        // Shrink the native overflow reference; grow the other profiles. Then
        // save again from the first saved bytes, retaining the same frame chain.
        foreach (int count in new[] { overflow ? 6 : 24, overflow ? 24 : 6 }) {
            OfficeRichTextParagraph template = Assert.Single(document.TextStories.Single(value => value.Id == story.Id).Paragraphs);
            OfficeRichTextRun run = Assert.Single(template.Runs);
            document.ReplaceStoryParagraphs(story.Id, new[] {
                new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(Text(count), run.FontSize, run.Color,
                    run.Bold, run.Italic, run.Underline, run.FontFamily) }, template.Alignment, template.LineHeight,
                    template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops)
            });
            var saved = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            Assert.False(saved.Report.HasErrors);
            PublisherNativeWritingTests.AssertPreservedStorage(source, saved.Value, "Quill/QuillSub/CONTENTS");
            Assert.Equal(Layout(source), Layout(saved.Value));
            document = PublisherDocument.Load(saved.Value);
            Assert.Equal(2, document.TextStories.Count);
            Assert.Equal(Text(count) + "\n", document.TextStories.Single(value => value.Id == story.Id).Text);
            Assert.Equal("Independent story\n", Assert.Single(document.TextStories, value => value.Id != story.Id).Text);
            PublisherTextFrame[] frames = Frames(document, story.Id);
            Assert.Equal(original.Select(Identity), frames.Select(Identity));
            Assert.Equal(count == 24, frames[2].HasOverflow);
            Assert.Equal(0, frames[0].TextStart);
            Assert.Equal(frames[0].TextLength, frames[1].TextStart);
            Assert.Equal(frames[1].TextStart + frames[1].TextLength, frames[2].TextStart);
            source = saved.Value;
        }
    }

    private static string Text(int count) => string.Concat(Enumerable.Repeat(Sentence, count)).TrimEnd();
    private static PublisherTextFrame[] Frames(PublisherDocument document, uint storyId) => document.Pages
        .SelectMany(page => page.TextFrames).Where(frame => frame.StoryId == storyId).OrderBy(frame => frame.Order).ToArray();
    private static (uint, uint, uint?, uint?, uint, double, double, double, double) Identity(PublisherTextFrame frame) =>
        (frame.Id, frame.PageId, frame.PreviousFrameId, frame.NextFrameId, frame.Order, frame.X, frame.Y, frame.Width, frame.Height);
    private static byte[] Layout(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] stream = compound!.Streams["Quill/QuillSub/CONTENTS"];
        var data = new PublisherBinaryData(stream, "Quill/QuillSub/CONTENTS");
        PublisherQuillChunk chunk = PublisherQuillDirectory.Read(data,
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None)).Single(value => value.Tag == "MCLD");
        return stream.Skip(chunk.Offset).Take(chunk.Length).ToArray();
    }
}
