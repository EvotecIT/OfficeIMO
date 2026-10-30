using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherParagraphTopologyWritingTests {
    [Theory]
    [InlineData("append")]
    [InlineData("remove")]
    [InlineData("allocate")]
    [InlineData("retire")]
    [InlineData("neighbor")]
    [InlineData("character-control")]
    public void Paragraph_count_changes_preserve_complete_stories_and_unrelated_native_storage(string variant) {
        string fixture = variant == "remove" ? "TextParagraphs/native-textparagraphs.pub"
            : variant == "retire" ? "TextFormattingMany/native-textformattingmany.pub"
            : variant == "neighbor" ? "TextStories/native-textstories.pub"
            : variant == "character-control" ? "TextCharacterPages/native-textcharacterpages.pub" : "Text/native-text.pub";
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory[] originals = document.TextStories.ToArray();
        PublisherTextStory story = originals[variant == "neighbor" ? 1 : 0];
        string[] text = variant == "allocate" ? Enumerable.Range(1, 240).Select(index => "New paragraph " + index).ToArray()
            : variant is "remove" or "retire" ? new[] { "Only remaining paragraph" }
            : variant == "character-control" ? new[] { "First new paragraph", "Second new paragraph" }
            : variant == "neighbor" ? new[] { "Middle one", "Middle two", "", "Middle four", "Middle five" }
            : new[] { "First new paragraph", "", "Third paragraph", "Fourth paragraph" };
        document.ReplaceStoryParagraphs(story.Id, Paragraphs(story, text, 18, true));
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument reopened = PublisherDocument.Load(result.Value);
        Assert.Equal(originals.Select(item => item.Id), reopened.TextStories.Select(item => item.Id));
        foreach (PublisherTextStory original in originals) {
            PublisherTextStory saved = reopened.TextStories.Single(item => item.Id == original.Id);
            Assert.Equal(original.Id == story.Id ? string.Join("\n", text) + "\n" : original.Text, saved.Text);
            if (original.Id == story.Id) {
                Assert.Equal(text.Length, saved.Paragraphs.Count);
                Assert.All(saved.Paragraphs, paragraph => {
                    OfficeRichTextRun run = Assert.Single(paragraph.Runs);
                    Assert.Equal(18, run.FontSize); Assert.True(run.Bold); Assert.Equal("Arial", run.FontFamily);
                });
            }
        }
        Assert.True(result.Report.HasLoss); Assert.False(result.Report.HasErrors);
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Quill/QuillSub/CONTENTS");
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("character")]
    public void Count_changes_refuse_ambiguous_native_style_mapping_even_when_projected_formatting_is_uniform(string ambiguity) {
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            if (ambiguity == "paragraph") {
                // The independent control has three FDPP ranges sharing a 4-byte OPL.
                PublisherInputContractTests.WriteUInt32(bytes, 18944 + 496, 6);
                bytes[18944 + 500] = 0xFE; bytes[18944 + 501] = 0x0A;
                bytes[18944 + 22] = 0xF0; bytes[18944 + 23] = 0x01;
            } else {
                // Split at the first native paragraph end, adding an unknown
                // Boolean field only to the following body's otherwise identical OPL.
                const int page = 19456;
                byte[] body = bytes.Skip(page + 370).Take(142).ToArray(), cr = bytes.Skip(page + 360).Take(10).ToArray();
                Array.Clear(bytes, page, 512);
                bytes[page] = 3; bytes[page + 2] = 1;
                PublisherInputContractTests.WriteUInt32(bytes, page + 8, 544);
                PublisherInputContractTests.WriteUInt32(bytes, page + 12, 574);
                PublisherInputContractTests.WriteUInt32(bytes, page + 16, 576);
                bytes[page + 20] = 200;
                bytes[page + 22] = 86; bytes[page + 23] = 1;
                bytes[page + 24] = 230; bytes[page + 25] = 1;
                Buffer.BlockCopy(body, 0, bytes, page + 200, body.Length);
                Buffer.BlockCopy(body, 0, bytes, page + 342, body.Length);
                PublisherInputContractTests.WriteUInt32(bytes, page + 342, 144);
                bytes[page + 484] = 0xFE; bytes[page + 485] = 0x0A;
                Buffer.BlockCopy(cr, 0, bytes, page + 486, cr.Length);
            }
        }, "NativeComparison/TextParagraphs/native-textparagraphs.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        PublisherTextStory story = Assert.Single(document.TextStories);
        Assert.All(story.Paragraphs, paragraph => { Assert.Equal(12, paragraph.Runs[0].FontSize); Assert.Equal("Arial", paragraph.Runs[0].FontFamily); });
        document.ReplaceStoryParagraphs(story.Id, Paragraphs(story, new[] { "New one", "New two", "New three", "New four" }, 12, false));
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
        Assert.Equal(4, document.TextStories[0].Paragraphs.Count);
    }

    [Fact]
    public void Directory_growth_refuses_to_overwrite_nonzero_unassigned_header_bytes() {
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => bytes[400] = 1,
            "NativeComparison/Text/native-text.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        PublisherTextStory story = Assert.Single(document.TextStories);
        document.ReplaceStoryParagraphs(story.Id, Paragraphs(story, Enumerable.Range(1, 240).Select(index => "Paragraph " + index).ToArray(), 12, false));
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
        Assert.Equal(240, document.TextStories[0].Paragraphs.Count);
    }

    [Fact]
    public void Character_and_paragraph_indexes_expand_together_without_losing_per_paragraph_styles() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"));
        PublisherTextStory story = Assert.Single(document.TextStories);
        var paragraphs = Enumerable.Range(0, 84).Select(index => Paragraph(story, "Paragraph " + index, 12 + Math.Min(index, 3) * 2, true)).ToArray();
        document.ReplaceStoryParagraphs(story.Id, paragraphs);
        PublisherDocument saved = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        PublisherTextStory current = Assert.Single(saved.TextStories);
        Assert.Equal(string.Join("\n", Enumerable.Range(0, 84).Select(index => "Paragraph " + index)) + "\n", current.Text);
        Assert.Equal(Enumerable.Range(0, 84).Select(index => (double)(12 + Math.Min(index, 3) * 2)), current.Paragraphs.Select(paragraph => paragraph.Runs[0].FontSize));
        foreach (string tag in new[] { "FDPC", "FDPP" })
            Assert.True(saved.SourceStyleCatalog.Stories[0].Uses.Where(use => use.Tag == tag).Select(use => use.ChunkId).Distinct().Count() > 1);
    }

    private static IReadOnlyList<OfficeRichTextParagraph> Paragraphs(PublisherTextStory source, string[] text, double size, bool bold) =>
        text.Select(value => Paragraph(source, value, size, bold)).ToArray();

    private static OfficeRichTextParagraph Paragraph(PublisherTextStory source, string text, double size, bool bold) {
        OfficeRichTextParagraph paragraph = source.Paragraphs[0]; OfficeRichTextRun run = Assert.Single(paragraph.Runs);
        return new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, run.Color, bold,
            run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle,
            run.StrikethroughStyle, run.Baseline, run.ParagraphIndent) }, paragraph.Alignment, paragraph.LineHeight,
            paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops);
    }
}
