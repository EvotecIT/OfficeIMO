using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTextStyleTests {
    [Fact]
    public void Referenced_native_styles_supply_paragraph_values_without_overriding_direct_formatting() {
        PublisherDocument newsletter = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        OfficeRichTextParagraph body = newsletter.TextStories.Single(story => story.Id == 2).Paragraphs[0];
        Assert.Equal(3, body.Margins.Bottom, 6); Assert.Equal(8.5, body.Runs[0].FontSize);
        OfficeRichTextParagraph heading = newsletter.TextStories.Single(story => story.Id == 15).Paragraphs[0];
        Assert.Equal(OfficeTextAlignment.Center, heading.Alignment); Assert.Equal(9, heading.Margins.Bottom);
        PublisherDocument brochure = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleBrochure.pub"));
        OfficeRichTextParagraph direct = brochure.TextStories.Single(story => story.Id == 13).Paragraphs
            .Single(paragraph => Text(paragraph).StartsWith("The key concept of Cubism"));
        Assert.Equal(0, direct.Margins.Bottom); Assert.Equal(9.75, direct.Runs[0].FontSize); Assert.Equal("Arial", direct.Runs[0].FontFamily);
    }

    [Fact]
    public void Native_bullets_remain_separate_from_source_text_and_keep_hanging_indentation() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherTextStory story = document.TextStories.Single(item => item.Id == 15);
        OfficeRichTextParagraph[] entries = story.Paragraphs.Where(paragraph => paragraph.Label != null).ToArray();
        Assert.Equal(3, entries.Length);
        Assert.All(entries, paragraph => {
            Assert.Equal("•", paragraph.Label!.Run.Text); Assert.Equal(paragraph.Runs[0].FontFamily, paragraph.Label.Run.FontFamily);
            Assert.Equal(0, paragraph.Label.Position); Assert.Equal(10.8, paragraph.Label.TextPosition);
            Assert.Equal(10.8, paragraph.Indent.ContinuationLineOffset); Assert.Equal(0, paragraph.Indent.FirstLineOffset);
        });
        Assert.DoesNotContain("•", story.Text);
        Assert.Contains("•", document.ToSvg(0));
    }

    [Fact]
    public void Native_tab_array_keeps_its_measured_position_in_source_paragraphs() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        OfficeRichTextParagraph paragraph = document.TextStories.Single(story => story.Id == 64).Paragraphs
            .Single(item => Text(item).StartsWith("Write out your 6 and 7 times tables"));
        OfficeTextTabStop stop = Assert.Single(paragraph.TabStops!.Stops);
        Assert.Equal(13.5, stop.Position); Assert.Equal(OfficeTextTabAlignment.Left, stop.Alignment);
        Assert.Equal(13.5, paragraph.Label!.TextPosition);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TAB_LAYOUT_APPROXIMATED");
    }

    [Fact]
    public void Malformed_native_tab_positions_and_counts_fail_before_rendering() {
        // Offsets are in the provenance-pinned newsletter's Quill stream.
        byte[] negative = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            Assert.Equal(171450U, BitConverter.ToUInt32(bytes, 30180));
            PublisherInputContractTests.WriteUInt32(bytes, 30180, uint.MaxValue);
        }, "SampleNewsletter.pub");
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(negative));
        byte[] mismatch = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            Assert.Equal(1, BitConverter.ToUInt16(bytes, 30164)); bytes[30164] = 2;
        }, "SampleNewsletter.pub");
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(mismatch));
    }

    [Fact]
    public void Unresolved_native_style_reference_retains_text_with_explicit_fallback_evidence() {
        byte[] input = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            Assert.Equal(2, BitConverter.ToUInt16(bytes, 29182)); bytes[29182] = 255; bytes[29183] = 255;
        }, "SampleNewsletter.pub");
        PublisherDocument document = PublisherDocument.Load(input);
        Assert.Contains(document.TextStories, story => story.Text.Contains("Living and Learning"));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TEXT_STYLE_REFERENCE_UNRESOLVED");
    }

    private static string Text(OfficeRichTextParagraph paragraph) => string.Concat(paragraph.Runs.Select(run => run.Text));
}
