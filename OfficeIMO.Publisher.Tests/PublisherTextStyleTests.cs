using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTextStyleTests {
    [Fact]
    public void Referenced_native_styles_supply_paragraph_values_without_overriding_direct_formatting() {
        PublisherDocument newsletter = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        OfficeRichTextParagraph body = newsletter.TextStories.Single(story => story.Id == 2).Paragraphs[0];
        Assert.Equal(3, body.Margins.Bottom, 6); Assert.Equal(8.5, body.Runs[0].FontSize);
        // Source coordinates and selectors were inspected independently in the
        // original corpus. Whole indexed entries include opaque OPL suffixes.
        PublisherQuillStyleCatalog catalog = newsletter.SourceStyleCatalog;
        Assert.Equal(25, catalog.Definitions.Count);
        PublisherQuillStyleDefinition pair = catalog.Definitions[2];
        Assert.Equal(2, pair.Ordinal); Assert.Equal(3D, pair.Paragraph.After); Assert.Null(pair.Paragraph.Alignment);
        Assert.Equal(new PublisherQuillSourceSpan(20352, 278), pair.CharacterEntry);
        Assert.Equal(new PublisherQuillSourceSpan(20354, 274), pair.CharacterDefinition);
        Assert.Equal(new PublisherQuillSourceSpan(20630, 18), pair.ParagraphEntry);
        Assert.Equal(new PublisherQuillSourceSpan(20632, 14), pair.ParagraphDefinition);
        PublisherQuillStoryStyles source = catalog.Stories.Single(story => story.StoryId == 2);
        Assert.Equal((0, 0, 867, 512, 10, 0),
            (source.StoryOrdinal, source.NativeStart, source.NativeLength, source.TextByteOffset, source.CarriageReturns, source.NullCharacters));
        PublisherQuillStyleUse first = source.Uses.First(use => use.Tag == "FDPP");
        Assert.Equal((0, 182, (uint?)2), (first.NativeStart, first.NativeLength, first.SelectedStyleOrdinal));
        Assert.Equal(new PublisherQuillSourceSpan(29176, 8), first.DirectDefinition);
        Assert.All(source.Uses.Where(use => use.Tag == "FDPP"), use => Assert.Equal(2U, use.SelectedStyleOrdinal));
        PublisherQuillStyleUse character = Assert.Single(source.Uses, use => use.Tag == "FDPC");
        Assert.Equal((0, 0, 0, 867), (character.ChunkId, character.RangeIndex, character.NativeStart, character.NativeLength));
        Assert.Equal(new PublisherQuillSourceSpan(30678, 42), character.DirectDefinition);
        Assert.Same(pair, source.Definitions.Single(definition => definition.Ordinal == 2));
        OfficeRichTextParagraph heading = newsletter.TextStories.Single(story => story.Id == 15).Paragraphs[0];
        Assert.Equal(OfficeTextAlignment.Center, heading.Alignment); Assert.Equal(9, heading.Margins.Bottom);
        Assert.Equal(10D, catalog.Definitions[5].Paragraph.LabelSize);
        PublisherQuillStoryStyles recipients = catalog.Stories.Single(story => story.StoryId == 15);
        PublisherQuillStyleUse recipient = recipients.Uses.Single(use => use.Tag == "FDPP" && use.NativeStart == 1468);
        Assert.Equal(5U, recipient.SelectedStyleOrdinal);
        Assert.Equal(new PublisherQuillSourceSpan(29110, 14), recipient.DirectDefinition);
        PublisherDocument brochure = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleBrochure.pub"));
        OfficeRichTextParagraph direct = brochure.TextStories.Single(story => story.Id == 13).Paragraphs
            .Single(paragraph => Text(paragraph).StartsWith("The key concept of Cubism"));
        Assert.Equal(0, direct.Margins.Bottom); Assert.Equal(9.75, direct.Runs[0].FontSize); Assert.Equal("Arial", direct.Runs[0].FontFamily);
        PublisherQuillStyleDefinition fallback = brochure.SourceStyleCatalog.Definitions[0];
        Assert.Equal(10D, fallback.Character.Size); Assert.Null(fallback.Paragraph.After);
        Assert.Equal(new PublisherQuillSourceSpan(8716, 8), fallback.ParagraphEntry);
        Assert.Equal(new PublisherQuillSourceSpan(8718, 4), fallback.ParagraphDefinition);
        PublisherQuillStoryStyles brochureSource = brochure.SourceStyleCatalog.Stories.Single(story => story.StoryId == 13);
        PublisherQuillStyleUse absent = brochureSource.Uses.Single(use => use.Tag == "FDPP" && use.NativeStart == 1315);
        Assert.Null(absent.SelectedStyleOrdinal); Assert.Equal(137, absent.NativeLength);
        Assert.Equal(new PublisherQuillSourceSpan(13694, 4), absent.DirectDefinition);
        PublisherQuillStyleUse directCharacter = brochureSource.Uses.Single(use => use.Tag == "FDPC" && use.NativeStart == 1315);
        Assert.Equal((1, 2, 533), (directCharacter.ChunkId, directCharacter.RangeIndex, directCharacter.NativeLength));
        Assert.Equal(new PublisherQuillSourceSpan(15144, 114), directCharacter.DirectDefinition);
    }

    [Fact]
    public void Native_bullets_remain_separate_from_source_text_and_keep_hanging_indentation() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherTextStory story = document.TextStories.Single(item => item.Id == 15);
        OfficeRichTextParagraph[] entries = story.Paragraphs.Where(paragraph => paragraph.Label != null).ToArray();
        Assert.Equal(3, entries.Length);
        Assert.All(entries, paragraph => {
            Assert.Equal("•", paragraph.Label!.Run.Text); Assert.Equal(paragraph.Runs[0].FontFamily, paragraph.Label.Run.FontFamily);
            // Native FDPP explicitly sets label size zero over pair five's
            // ten-point default; the label therefore follows the 9.2-point run.
            Assert.Equal(9.2, paragraph.Runs[0].FontSize); Assert.Equal(9.2, paragraph.Label.Run.FontSize);
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

    [Theory]
    [InlineData("SampleNewsletter.pub", "Jay Adams")]
    [InlineData("SampleBrochure.pub", "Saturday")]
    public void Native_tabs_without_arrays_use_a_measured_default_grid(string file, string prefix) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(file));
        OfficeRichTextParagraph paragraph = document.TextStories.SelectMany(story => story.Paragraphs)
            .Single(item => Text(item).StartsWith(prefix) && Text(item).Contains("\t"));
        Assert.NotNull(paragraph.TabStops); Assert.Empty(paragraph.TabStops!.Stops);
        Assert.Equal(36, paragraph.TabStops.DefaultInterval); Assert.Equal(-paragraph.Margins.Left, paragraph.TabStops.Origin);
        string[] fields = Text(paragraph).Split('\t'); Assert.Equal(2, fields.Length);
        OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 400, 400,
            (value, size, family, style) => value?.Length ?? 0);
        Assert.Equal(36 + fields[1].Length + paragraph.Margins.Right, layout.Width, 6);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, finding => finding.Code == "PUB_TAB_LAYOUT_APPROXIMATED"
            && finding.LossKind == OfficeConversionLossKind.Approximation);
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
