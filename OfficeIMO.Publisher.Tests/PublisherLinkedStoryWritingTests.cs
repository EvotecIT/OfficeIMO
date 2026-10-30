using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherLinkedStoryWritingTests {
    private const string Fixture = "NativeComparison/Newsletter/native-newsletter.pub";

    [Theory]
    [InlineData("grow")]
    [InlineData("shrink")]
    [InlineData("allocate")]
    [InlineData("retire")]
    [InlineData("styles")]
    [InlineData("neighbor")]
    public void Complete_linked_story_edits_preserve_chain_ownership_and_unrelated_native_storage(string variant) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(Fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory original = document.TextStories.Single(story => story.Id == 3);
        PublisherTextFrame[] frames = Frames(document);
        Assert.Equal(new uint[] { 297, 298 }, frames.Select(frame => frame.Id));
        string[] text = variant == "allocate" ? Enumerable.Range(1, 240).Select(index => "Linked paragraph " + index).ToArray()
            : variant == "retire" ? new[] { "Only remaining linked paragraph" }
            : original.Paragraphs.Select((paragraph, index) => variant == "shrink" ? "Short " + index
                : "Linked section " + index + ". " + new string('x', 180)).ToArray();
        OfficeRichTextParagraph template = original.Paragraphs[0]; OfficeRichTextRun run = Assert.Single(template.Runs);
        document.ReplaceStoryParagraphs(original.Id, text.Select(value => Paragraph(template, value,
            variant == "styles" ? 18 : run.FontSize, variant == "styles" || run.Bold)).ToArray());
        if (variant == "neighbor") {
            OfficeRichTextParagraph title = document.TextStories.Single(story => story.Id == 2).Paragraphs[0];
            document.ReplaceStoryParagraphs(2, new[] { Paragraph(title, "Updated newsletter title", title.Runs[0].FontSize, title.Runs[0].Bold) });
        }
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.False(result.Report.HasErrors); Assert.True(result.Report.HasLoss);
        PublisherDocument reopened = PublisherDocument.Load(result.Value);
        Assert.Equal(document.TextStories.Select(story => story.Text), reopened.TextStories.Select(story => story.Text));
        Assert.Equal(text.Length, reopened.TextStories.Single(story => story.Id == 3).Paragraphs.Count);
        PublisherTextFrame[] savedFrames = Frames(reopened);
        Assert.Equal(frames.Select(Identity), savedFrames.Select(Identity));
        Assert.Equal(document.Pages.Select(page => page.Id), reopened.Pages.Select(page => page.Id));
        Assert.All(reopened.TextStories.Single(story => story.Id == 3).Paragraphs, paragraph => {
            OfficeRichTextRun saved = Assert.Single(paragraph.Runs);
            Assert.Equal(variant == "styles" ? 18 : run.FontSize, saved.FontSize);
            Assert.Equal(variant == "styles" || run.Bold, saved.Bold); Assert.Equal(run.FontFamily, saved.FontFamily);
        });
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Quill/QuillSub/CONTENTS");
    }

    [Fact]
    public void Unresolved_native_chain_remains_exact_copyable_but_refuses_changed_story_even_when_loss_is_allowed() {
        byte[] source = PublisherInputContractTests.Mutate("Contents", bytes => {
            var data = new PublisherBinaryData(bytes, "Contents");
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            var blocks = new PublisherBlockReader(data, context);
            PublisherContents publication = new PublisherContentsReader(data, context).Read();
            PublisherContentChunk document = publication.Chunks.Values.Single(chunk => chunk.Kind == 0x44);
            PublisherBlock catalog = PublisherContentsReader.Field(blocks.Chunk(document.Offset, document.End), 0x20)!.Value;
            bytes[catalog.Offset - 2] = 0xFE;
        }, Fixture);
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_CHAIN_UNRESOLVED");
        OfficeRichTextParagraph template = document.TextStories.Single(story => story.Id == 3).Paragraphs[0];
        document.ReplaceStoryParagraphs(3, new[] { Paragraph(template, "Changed unresolved story", template.Runs[0].FontSize, template.Runs[0].Bold) });
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(Assert.IsType<PublisherSaveReport>(error.Report).FidelityDiagnostics,
            item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED" && item.LossKind == OfficeConversionLossKind.Failure);
        Assert.Equal("Changed unresolved story\n", document.TextStories.Single(story => story.Id == 3).Text);
    }

    private static OfficeRichTextParagraph Paragraph(OfficeRichTextParagraph template, string text, double size, bool bold) {
        OfficeRichTextRun run = template.Runs[0];
        return new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, run.Color, bold,
            run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor,
            run.UnderlineStyle, run.StrikethroughStyle, run.Baseline, run.ParagraphIndent)
        }, template.Alignment, template.LineHeight, template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops);
    }

    private static PublisherTextFrame[] Frames(PublisherDocument document) => document.Pages.SelectMany(page => page.TextFrames)
        .Where(frame => frame.StoryId == 3).OrderBy(frame => frame.Order).ToArray();
    private static (uint, uint, uint?, uint?, uint, double, double, double, double) Identity(PublisherTextFrame frame) =>
        (frame.Id, frame.PageId, frame.PreviousFrameId, frame.NextFrameId, frame.Order, frame.X, frame.Y, frame.Width, frame.Height);
}
