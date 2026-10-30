using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherStyleProvenanceTests {
    [Fact]
    public void Publication_without_TEXT_retains_its_indexed_style_catalog() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Blank/native-blank.pub"));
        Assert.Empty(document.TextStories);
        PublisherQuillStyleCatalog catalog = document.SourceStyleCatalog;
        Assert.Empty(catalog.Stories); Assert.Equal(36, catalog.Definitions.Count);
        // The independent original-byte reference records no TEXT descriptor,
        // and these three indexed STSH envelopes still belong to that source.
        Assert.Equal(new[] { 0, 1, 2 }, catalog.Libraries.Select(library => library.Id));
        Assert.Equal(new[] { 36, 72, 72 }, catalog.Libraries.Select(library => library.Entries.Count));
        Assert.Equal(new[] { new PublisherQuillSourceSpan(512, 380), new PublisherQuillSourceSpan(892, 16442),
            new PublisherQuillSourceSpan(17334, 1460) }, catalog.Libraries.Select(library => library.Source));
        Assert.Contains(document.ReadReport.FidelityDiagnostics, finding => finding.Code == "PUB_NAMED_STYLE_METADATA_UNASSESSED");
    }

    [Fact]
    public void Native_style_ranges_keep_CR_and_NUL_coordinates_separate_from_normalized_text() {
        byte[] input = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            // Original newsletter story two begins with T at TEXT byte 512.
            Assert.Equal(0x54, bytes[512]); Assert.Equal(0, bytes[513]); bytes[512] = 0;
        }, "SampleNewsletter.pub");
        PublisherDocument document = PublisherDocument.Load(input);
        PublisherTextStory story = document.TextStories.Single(item => item.Id == 2);
        PublisherQuillStoryStyles source = document.SourceStyleCatalog.Stories.Single(item => item.StoryId == 2);
        Assert.Equal((0, 0, 867, 512, 10, 1),
            (source.StoryOrdinal, source.NativeStart, source.NativeLength, source.TextByteOffset, source.CarriageReturns, source.NullCharacters));
        Assert.Equal(866, story.Text.Length); Assert.Equal(180, story.Text.IndexOf('\n'));
        PublisherQuillStyleUse[] paragraphs = source.Uses.Where(use => use.Tag == "FDPP").ToArray();
        Assert.Equal((0, 182), (paragraphs[0].NativeStart, paragraphs[0].NativeLength));
        Assert.Equal(182, paragraphs[1].NativeStart);
        Assert.Equal(new PublisherQuillSourceSpan(29176, 8), paragraphs[0].DirectDefinition);
        PublisherQuillStyleUse characters = Assert.Single(source.Uses, use => use.Tag == "FDPC");
        Assert.Equal((0, 867), (characters.NativeStart, characters.NativeLength));
        Assert.Equal(new PublisherQuillSourceSpan(30678, 42), characters.DirectDefinition);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Story_edits_retain_original_style_evidence_and_refuse_native_bytes(bool formattingOnly) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherQuillStyleCatalog catalog = document.SourceStyleCatalog;
        PublisherQuillStoryStyles source = catalog.Stories.Single(story => story.StoryId == 2);
        PublisherTextStory original = document.TextStories.Single(story => story.Id == 2);
        OfficeRichTextParagraph[] paragraphs;
        if (formattingOnly) {
            paragraphs = original.Paragraphs.ToArray();
            OfficeRichTextParagraph first = paragraphs[0];
            OfficeTextAlignment changed = first.Alignment == OfficeTextAlignment.Left ? OfficeTextAlignment.Right : OfficeTextAlignment.Left;
            paragraphs[0] = new OfficeRichTextParagraph(first.Runs, changed, first.LineHeight, first.Margins,
                first.Indent, first.LineHeightFactor).WithTabStops(first.TabStops);
        } else paragraphs = new[] { new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Edited", 12, OfficeColor.Black, fontFamily: "Arial") }) };
        document.ReplaceStoryParagraphs(2, paragraphs);
        string edited = document.TextStories.Single(story => story.Id == 2).Text;
        if (formattingOnly) Assert.Equal(original.Text, edited); else Assert.NotEqual(original.Text, edited);
        Assert.Same(catalog, document.SourceStyleCatalog);
        PublisherNativePreparation prepared = document.PrepareNativeBytes(64L * 1024 * 1024);
        Assert.False(prepared.IsPrepared); Assert.Equal(0, prepared.ByteLength);
        Assert.Throws<NotSupportedException>(() => prepared.CopyBytes());
        Assert.Contains(prepared.FidelityDiagnostics, finding => finding.Code == "PUB_NATIVE_EDIT_UNSUPPORTED"
            && finding.Location == "Quill" && finding.LossKind == OfficeConversionLossKind.Failure);
        PublisherQuillStoryStyles dependency = Assert.Single(prepared.StyleDependencies);
        Assert.Same(source, dependency);
        Assert.Equal((2U, 0, 0, 867), (dependency.StoryId, dependency.StoryOrdinal, dependency.NativeStart, dependency.NativeLength));
        Assert.Contains(dependency.Definitions, definition => definition.Ordinal == 2
            && definition.ParagraphDefinition == new PublisherQuillSourceSpan(20632, 14));
        Assert.Contains(dependency.Uses, use => use.Tag == "FDPC" && use.DirectDefinition == new PublisherQuillSourceSpan(30678, 42));
        Assert.Equal(edited, document.TextStories.Single(story => story.Id == 2).Text);
    }

    [Fact]
    public void Opaque_style_library_index_outside_its_payload_is_rejected_without_TEXT() {
        byte[] input = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            Assert.Equal(144U, BitConverter.ToUInt32(bytes, 532));
            PublisherInputContractTests.WriteUInt32(bytes, 532, uint.MaxValue);
        }, "NativeComparison/Blank/native-blank.pub");
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input));
    }
}
