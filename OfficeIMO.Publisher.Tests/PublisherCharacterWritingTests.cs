using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherCharacterWritingTests {
    [Theory]
    [InlineData("italic")]
    [InlineData("strike")]
    [InlineData("words")]
    [InlineData("double")]
    [InlineData("dotted")]
    [InlineData("dashed")]
    [InlineData("wavy")]
    [InlineData("super")]
    [InlineData("sub")]
    [InlineData("font-existing")]
    [InlineData("font-new")]
    [InlineData("color-existing")]
    [InlineData("color-new")]
    [InlineData("combined")]
    [InlineData("chained-remove")]
    [InlineData("chained-pattern")]
    [InlineData("paragraphs")]
    [InlineData("linked")]
    public void Native_character_edits_preserve_complete_stories_attributes_and_unrelated_storage(string variant) {
        string fixture = variant is "color-existing" or "chained-remove" ? "TextCharacterAttributes/native-textcharacterattributes.pub"
            : variant == "chained-pattern" ? "TextUnderlinePatterns/native-textunderlinepatterns.pub"
            : variant == "linked" ? "Newsletter/native-newsletter.pub" : "Text/native-text.pub";
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        uint id = variant == "color-existing" ? 12U : variant == "chained-remove" ? 13U
            : variant == "chained-pattern" ? 19U : variant == "linked" ? 3U : 2U;
        PublisherTextStory story = document.TextStories.Single(item => item.Id == id);
        OfficeRichTextParagraph template = story.Paragraphs[0];
        OfficeRichTextRun original = Assert.Single(template.Runs);
        OfficeTextDecorationStyle underline = variant is "double" or "combined" or "chained-pattern" ? OfficeTextDecorationStyle.Double
            : variant == "words" ? OfficeTextDecorationStyle.Words : variant == "dotted" ? OfficeTextDecorationStyle.Dotted
            : variant == "dashed" ? OfficeTextDecorationStyle.Dashed : variant == "wavy" ? OfficeTextDecorationStyle.Wavy : OfficeTextDecorationStyle.None;
        string family = variant == "font-existing" ? "Cambria" : variant is "font-new" or "combined" or "paragraphs" or "linked" ? "Consolas" : "Arial";
        OfficeColor color = variant is "color-existing" or "combined" or "linked" ? OfficeColor.Blue
            : variant is "color-new" or "paragraphs" ? OfficeColor.FromRgb(14, 93, 151) : OfficeColor.Black;
        OfficeTextBaseline baseline = variant is "super" or "combined" ? OfficeTextBaseline.Superscript
            : variant == "sub" ? OfficeTextBaseline.Subscript : OfficeTextBaseline.Normal;
        double size = variant == "combined" ? 18 : 12;
        bool bold = variant == "combined", italic = variant is "italic" or "combined" or "paragraphs" or "linked";
        bool strike = variant is "strike" or "combined";
        string[] text = variant == "paragraphs" ? new[] { "First styled paragraph", "Second styled paragraph", "Third styled paragraph" }
            : variant == "linked" ? Enumerable.Range(1, story.Paragraphs.Count).Select(index => "Linked styled section " + index + ". " + new string('x', 120)).ToArray()
            : new[] { "Changed text for " + variant };
        document.ReplaceStoryParagraphs(id, text.Select(value => new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun(value, size, color, bold, italic, fontFamily: family, strikethrough: strike,
                underlineStyle: underline, baseline: baseline)
        }, template.Alignment, template.LineHeight, template.Margins, template.Indent, template.LineHeightFactor)
            .WithTabStops(template.TabStops)).ToArray());
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.False(result.Report.HasErrors); Assert.True(result.Report.HasLoss);
        PublisherDocument saved = PublisherDocument.Load(result.Value);
        Assert.Equal(document.TextStories.Select(item => item.Text), saved.TextStories.Select(item => item.Text));
        Assert.All(saved.TextStories.Single(item => item.Id == id).Paragraphs, paragraph => {
            OfficeRichTextRun run = Assert.Single(paragraph.Runs);
            Assert.Equal((size, color, family, bold, italic, strike, underline, baseline),
                (run.FontSize, run.Color, run.FontFamily, run.Bold, run.Italic, run.Strikethrough, run.UnderlineStyle, run.Baseline));
        });
        Assert.Equal(Frames(document), Frames(saved));
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Quill/QuillSub/CONTENTS");
        Assert.Equal(source, PublisherDocument.Load(source).ToPubBytes());
    }

    [Theory]
    [InlineData("double-strike")]
    [InlineData("font-length")]
    [InlineData("alpha")]
    public void Native_unrepresentable_character_requests_remain_refused_even_when_loss_is_allowed(string variant) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"));
        document.ReplaceStoryParagraphs(2, new[] { new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("Unsupported request", 12, variant == "alpha" ? OfficeColor.FromRgba(255, 0, 0, 128) : OfficeColor.Black,
                fontFamily: variant == "font-length" ? new string('x', 33) : "Arial",
                strikethroughStyle: variant == "double-strike" ? OfficeTextDecorationStyle.Double : OfficeTextDecorationStyle.None)
        }) });
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
        Assert.Contains("Unsupported request", document.TextStories[0].Text);
    }

    [Fact]
    public void Font_growth_refuses_a_contradictory_index_while_unchanged_copying_retains_the_inert_source() {
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            PublisherQuillChunk font = PublisherQuillDirectory.Read(new PublisherBinaryData(bytes, "Quill"), context).Single(chunk => chunk.Tag == "FONT");
            bytes[font.Offset + 20] ^= 1;
        }, "NativeComparison/Text/native-text.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        document.ReplaceStoryParagraphs(2, new[] { new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("New font", 12, OfficeColor.Black, fontFamily: "Consolas")
        }) });
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Chained_directory_edits_preserve_unknown_padding_and_refuse_unknown_header_metadata(bool unknownHeader) {
        int marker = 0;
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            int secondary = checked((int)BitConverter.ToUInt32(bytes, 28));
            marker = secondary + (unknownHeader ? 0 : 100);
            bytes[marker] = unknownHeader ? (byte)0xF9 : (byte)0x5A;
        }, "NativeComparison/TextCharacterAttributes/native-textcharacterattributes.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        OfficeRichTextParagraph template = document.TextStories.Single(story => story.Id == 2).Paragraphs[0];
        document.ReplaceStoryParagraphs(2, new[] { new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("Keep padding", 12, OfficeColor.Black, italic: true, fontFamily: "Arial")
        }, template.Alignment, template.LineHeight, template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops) });
        if (unknownHeader) {
            Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        } else {
            byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            Assert.True(OfficeCompoundFileReader.TryRead(output, out OfficeCompoundFile? compound, out string? error), error);
            Assert.Equal(0x5A, compound!.Streams["Quill/QuillSub/CONTENTS"][marker]);
            Assert.True(PublisherDocument.Load(output).TextStories.Single(story => story.Id == 2).Paragraphs[0].Runs[0].Italic);
        }
    }

    private static IEnumerable<(uint, uint, uint, uint?, uint?, double, double, double, double)> Frames(PublisherDocument document) =>
        document.Pages.SelectMany(page => page.TextFrames).Select(frame =>
            (frame.Id, frame.StoryId, frame.PageId, frame.PreviousFrameId, frame.NextFrameId, frame.X, frame.Y, frame.Width, frame.Height));
}
