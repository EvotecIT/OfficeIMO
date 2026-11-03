using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherMultiParagraphCreationTests {
    private const long MaximumBytes = 64L * 1024 * 1024;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Ordinary_multi_paragraph_creation_preserves_explicit_empty_typing_styles_and_complete_overflow(bool replace) {
        PublisherDocument document = AuthoredDocument();
        OfficeRichTextParagraph[] paragraphs = Paragraphs();
        PublisherTextFrame frame = document.AddTextFrame(document.Pages[0].Id, 72, 72, 216, 72,
            replace ? new[] { new OfficeRichTextParagraph(new[] { Run("Placeholder", 12, false) }, lineHeight: 18) } : paragraphs);
        if (replace) document.ReplaceStoryParagraphs(frame.StoryId, paragraphs);
        document.SetStoryAutomaticHyphenation(frame.StoryId, false);
        PublisherTextStory authored = Assert.Single(document.TextStories);
        Assert.Equal(CompleteText(paragraphs), authored.Text);
        Assert.False(authored.HasFinalParagraphTerminator);
        Assert.True(Assert.Single(document.Pages[0].TextFrames).HasOverflow);
        PublisherNativeCreationPlan plan = PublisherNativeCreationProjection.FrameLayouts(document);
        byte[] generated = PublisherNativeCreationCodec.Create(plan, new PublisherReadOptions(), MaximumBytes, CancellationToken.None);
        PublisherDocument decoded = PublisherDocument.Load(generated);
        PublisherTextStory story = Assert.Single(decoded.TextStories);
        Assert.Equal(CompleteText(paragraphs) + "\n", story.Text);
        Assert.True(story.HasFinalParagraphTerminator);
        AssertParagraphs(paragraphs, story.Paragraphs);
        PublisherTextFrame restored = Assert.Single(Assert.Single(decoded.Pages).TextFrames);
        Assert.True(restored.HasOverflow);
        Assert.Equal(0, restored.TextStart);
        Assert.True(restored.TextLength < story.Text.Length);
        AssertCarriageReturnTyping(generated, paragraphs);
    }

    [Fact]
    public void Linked_multi_paragraph_creation_keeps_one_complete_story_in_chain_order_across_reordered_pages() {
        PublisherDocument document = AuthoredDocument();
        uint masterId = Assert.Single(document.MasterPages).Id;
        document.AddPage(masterId); document.AddPage(masterId);
        OfficeRichTextParagraph[] paragraphs = Paragraphs();
        PublisherTextFrame first = document.AddTextFrame(document.Pages[0].Id, 72, 72, 216, 72, paragraphs);
        document.SetStoryAutomaticHyphenation(first.StoryId, false);
        PublisherTextFrame middle = document.AddLinkedTextFrame(first.Id, document.Pages[2].Id, 72, 144, 216, 72);
        document.AddLinkedTextFrame(middle.Id, document.Pages[1].Id, 72, 216, 360, 108);
        Assert.Equal(CompleteText(paragraphs), Assert.Single(document.TextStories).Text);
        Assert.True(Assert.Single(document.Pages[1].TextFrames).HasOverflow);
        PublisherNativeCreationPlan plan = PublisherNativeCreationProjection.LinkedStories(document);
        byte[] generated = PublisherNativeCreationCodec.Create(plan, new PublisherReadOptions(), MaximumBytes, CancellationToken.None);
        PublisherDocument decoded = PublisherDocument.Load(generated);
        PublisherTextStory story = Assert.Single(decoded.TextStories);
        Assert.Equal(CompleteText(paragraphs) + "\n", story.Text);
        AssertParagraphs(paragraphs, story.Paragraphs);
        PublisherTextFrame[] frames = decoded.Pages.SelectMany(page => page.TextFrames).OrderBy(frame => frame.Order).ToArray();
        Assert.Equal(3, frames.Length);
        Assert.All(frames, frame => Assert.Equal(story.Id, frame.StoryId));
        Assert.Equal(new uint[] { 0, 1, 2 }, frames.Select(frame => frame.Order));
        var pageIndices = decoded.Pages.Select((page, index) => (page.Id, index)).ToDictionary(item => item.Id, item => item.index);
        Assert.Equal(new[] { 0, 2, 1 }, frames.Select(frame => pageIndices[frame.PageId]));
        Assert.Null(frames[0].PreviousFrameId); Assert.Null(frames[2].NextFrameId);
        Assert.Equal(0, frames[0].TextStart);
        Assert.All(frames, frame => { Assert.True(frame.TextStart.HasValue); Assert.True(frame.TextLength.HasValue); });
        for (int index = 1; index < frames.Length; index++) {
            Assert.Equal(frames[index - 1].Id, frames[index].PreviousFrameId);
            Assert.Equal(frames[index].Id, frames[index - 1].NextFrameId);
            Assert.Equal(frames[index - 1].TextStart + frames[index - 1].TextLength, frames[index].TextStart);
        }
        string assigned = string.Concat(frames.Select(frame => story.Text.Substring(frame.TextStart!.Value, frame.TextLength!.Value)));
        Assert.StartsWith(assigned, story.Text);
        Assert.True(assigned.Length < story.Text.Length);
        Assert.True(frames[2].HasOverflow);
    }

    [Theory]
    [InlineData("paragraphs")]
    [InlineData("runs")]
    public void Ordinary_creation_projection_refuses_missing_paragraphs_or_explicit_runs(string missing) {
        PublisherDocument document = AuthoredDocument();
        PublisherTextFrame frame = document.AddTextFrame(document.Pages[0].Id, 72, 72, 216, 72,
            new[] { new OfficeRichTextParagraph(new[] { Run("Placeholder", 12, false) }, lineHeight: 18) });
        OfficeRichTextParagraph[] requested = missing == "paragraphs" ? Array.Empty<OfficeRichTextParagraph>()
            : new[] { new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>(), lineHeight: 18) };
        document.ReplaceStoryParagraphs(frame.StoryId, requested);
        PublisherTextStory story = Assert.Single(document.TextStories);
        Assert.Equal(requested.Length, story.Paragraphs.Count);
        Assert.Empty(story.Text);
        if (missing == "runs") Assert.Empty(Assert.Single(story.Paragraphs).Runs);
        Assert.Throws<NotSupportedException>(() => PublisherNativeCreationProjection.FrameLayouts(document));
    }

    private static PublisherDocument AuthoredDocument() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        return document;
    }

    private static OfficeRichTextParagraph[] Paragraphs() => new[] {
        new OfficeRichTextParagraph(new[] { Run("Multi-paragraph heading. ", 16, true), Run("Body follows below.", 12, false) }, lineHeight: 22),
        new OfficeRichTextParagraph(new[] { Run(string.Empty, 14, true) }, lineHeight: 20),
        new OfficeRichTextParagraph(new[] {
            Run("Final paragraph: ", 12, true),
            Run(string.Concat(Enumerable.Repeat("Complete ordinary stories retain every paragraph and all overflow text. ", 18)).TrimEnd(), 12, false)
        }, lineHeight: 18)
    };

    private static OfficeRichTextRun Run(string text, double size, bool bold) =>
        new OfficeRichTextRun(text, size, OfficeColor.Black, bold, fontFamily: "Arial").WithLanguageId(1033);

    private static string CompleteText(IReadOnlyList<OfficeRichTextParagraph> paragraphs) =>
        string.Join("\n", paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));

    private static void AssertParagraphs(IReadOnlyList<OfficeRichTextParagraph> expected, IReadOnlyList<OfficeRichTextParagraph> actual) {
        Assert.Equal(expected.Count, actual.Count);
        for (int index = 0; index < expected.Count; index++) {
            Assert.Equal(expected[index].LineHeight, actual[index].LineHeight);
            Assert.Equal(expected[index].Runs.Select(run => (run.Text, run.FontSize, run.Bold, run.FontFamily, run.Color, run.LanguageId)),
                actual[index].Runs.Select(run => (run.Text, run.FontSize, run.Bold, run.FontFamily, run.Color, run.LanguageId)));
        }
    }

    /// <summary>Observes the produced CR characters and their resolved typing through existing Quill readers.</summary>
    private static void AssertCarriageReturnTyping(byte[] generated, IReadOnlyList<OfficeRichTextParagraph> paragraphs) {
        Assert.True(OfficeCompoundFileReader.TryRead(generated, out OfficeCompoundFile? compound, out string? error), error);
        var data = new PublisherBinaryData(compound!.Streams["Quill/QuillSub/CONTENTS"], "Quill/QuillSub/CONTENTS");
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        IReadOnlyList<PublisherQuillChunk> chunks = PublisherQuillDirectory.Read(data, context);
        PublisherQuillChunk text = Assert.Single(chunks, chunk => chunk.Tag == "TEXT");
        PublisherQuillStyles styles = new PublisherQuillStyleReader(data, context, chunks, Array.Empty<OfficeColor?>(), text).Read();
        int position = -1;
        foreach (OfficeRichTextParagraph paragraph in paragraphs) {
            position += paragraph.Runs.Sum(run => run.Text.Length) + 1;
            Assert.Equal('\r', (char)data.U16(text.Offset + position * 2));
            PublisherCharacterRange range = Assert.IsType<PublisherCharacterRange>(styles.CharacterAt(position));
            OfficeRichTextRun actual = styles.Run("\r", range.Style, styles.ParagraphAt(position));
            OfficeRichTextRun expected = paragraph.Runs[paragraph.Runs.Count - 1];
            Assert.Equal((expected.FontSize, expected.Bold, expected.FontFamily, expected.Color, expected.LanguageId),
                (actual.FontSize, actual.Bold, actual.FontFamily, actual.Color, actual.LanguageId));
        }
        Assert.Equal((position + 1) * 2, text.Length);
    }
}
