using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherMultipleStoryWritingTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(-1)]
    public void Independent_story_edits_retain_neighbor_text_and_native_story_order(int edited) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TextStories/native-textstories.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory[] stories = document.TextStories.ToArray();
        Assert.Equal(new uint[] { 2, 3, 4 }, stories.Select(story => story.Id));
        Assert.Equal(new[] { "First independent story\n", "Middle paragraph\n\nLast middle paragraph\n", "Last independent story\n" }, stories.Select(story => story.Text));
        var expected = stories.Select(story => story.Text).ToArray();
        for (int index = 0; index < stories.Length; index++) {
            if (edited != -1 && edited != index) continue;
            string[] text = index == 1 ? new[] { "Changed middle " + new string('i', 300), "Filled empty", "Tail \U0001F642" }
                : new[] { index == 0 ? "First grows " + new string('i', 300) : "Last shrinks" };
            document.ReplaceStoryParagraphs(stories[index].Id, Replace(stories[index], text, 18, true));
            expected[index] = string.Join("\n", text) + "\n";
        }
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.True(result.Report.HasLoss); Assert.False(result.Report.HasErrors);
        PublisherDocument reopened = PublisherDocument.Load(result.Value);
        Assert.Equal(new uint[] { 2, 3, 4 }, reopened.TextStories.Select(story => story.Id));
        Assert.Equal(expected, reopened.TextStories.Select(story => story.Text));
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Quill/QuillSub/CONTENTS");
        Assert.Equal(3, reopened.Pages[0].TextFrames.Count);
        Assert.Equal(source, File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TextStories/native-textstories.pub")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Existing_multiple_formatting_pages_cover_every_changed_paragraph_and_retained_overflow(bool shrink) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TextFormattingPages/native-textformattingpages.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory original = Assert.Single(document.TextStories);
        Assert.Equal(96, original.Paragraphs.Count);
        Assert.True(Assert.Single(document.Pages[0].TextFrames).HasOverflow);
        string[] text = Enumerable.Range(1, 96).Select(index => shrink ? "P" + index : "Changed paragraph " + index + new string('i', 24)).ToArray();
        document.ReplaceStoryParagraphs(original.Id, Replace(original, text, 18, true));
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherTextStory saved = Assert.Single(PublisherDocument.Load(result.Value).TextStories);
        Assert.Equal(string.Join("\n", text) + "\n", saved.Text);
        Assert.Equal(96, saved.Paragraphs.Count);
        Assert.All(saved.Paragraphs, paragraph => { Assert.Equal(18, paragraph.Runs[0].FontSize); Assert.True(paragraph.Runs[0].Bold); });
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Quill/QuillSub/CONTENTS");
        Assert.Throws<InvalidDataException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow, MaxOutputBytes = 1 }));
    }

    [Theory]
    [InlineData(20514)]
    [InlineData(20538)]
    public void Malformed_native_format_indexes_refuse_edits_while_exact_copying_remains_inert(int indexOffset) {
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => bytes[indexOffset + 16] ^= 1,
            "NativeComparison/TextStories/native-textstories.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        PublisherTextStory story = document.TextStories[0];
        document.ReplaceStoryParagraphs(story.Id, Replace(story, new[] { "Changed first story" }, 18, true));
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        var report = Assert.IsType<PublisherSaveReport>(error.Report);
        Assert.True(report.HasErrors); Assert.Equal(0, report.OutputByteCount);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Location == "Quill" && diagnostic.Code == "PUB_NATIVE_EDIT_UNSUPPORTED");
        Assert.Equal("Changed first story\n", document.TextStories[0].Text);
    }

    [Fact]
    public void Distinct_styles_allocate_chained_directories_without_losing_overflow_or_formatting() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TextFormattingPages/native-textformattingpages.pub"));
        PublisherTextStory source = Assert.Single(document.TextStories);
        var paragraphs = source.Paragraphs.Select((paragraph, index) => ReplaceParagraph(paragraph, paragraph.Runs[0].Text, 12 + index, true)).ToArray();
        document.ReplaceStoryParagraphs(source.Id, paragraphs);
        byte[] bytes = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument reopened = PublisherDocument.Load(bytes);
        Assert.Equal(source.Text, reopened.TextStories[0].Text);
        Assert.Equal(Enumerable.Range(12, 96).Select(size => (double)size), reopened.TextStories[0].Paragraphs.Select(paragraph => paragraph.Runs[0].FontSize));
        Assert.All(reopened.TextStories[0].Paragraphs, paragraph => Assert.True(paragraph.Runs[0].Bold));
        Assert.True(reopened.Pages[0].TextFrames[0].HasOverflow);
    }

    private static IReadOnlyList<OfficeRichTextParagraph> Replace(PublisherTextStory story, string[] text, double size, bool bold) =>
        story.Paragraphs.Select((paragraph, index) => ReplaceParagraph(paragraph, text[index], size, bold)).ToArray();

    private static OfficeRichTextParagraph ReplaceParagraph(OfficeRichTextParagraph paragraph, string text, double size, bool bold) {
            OfficeRichTextRun run = Assert.Single(paragraph.Runs);
            return new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, run.Color, bold,
                run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle,
                run.StrikethroughStyle, run.Baseline, run.ParagraphIndent) }, paragraph.Alignment, paragraph.LineHeight,
                paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops);
    }
}
