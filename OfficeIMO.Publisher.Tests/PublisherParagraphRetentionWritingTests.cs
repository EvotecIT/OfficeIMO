using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherParagraphRetentionWritingTests {
    private const string Fixture = "NativeComparison/TextNonuniformParagraphs/native-textnonuniformparagraphs.pub";

    [Theory]
    [InlineData("drop-first", false)]
    [InlineData("drop-middle", false)]
    [InlineData("drop-final", false)]
    [InlineData("reorder", false)]
    [InlineData("duplicate", false)]
    [InlineData("reorder", true)]
    public void Retained_nonuniform_paragraphs_keep_their_native_definitions_and_unrelated_stories(string operation, bool detached) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(Fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory original = document.TextStories[0];
        int[] order = operation switch {
            "drop-first" => new[] { 1, 2, 3 }, "drop-middle" => new[] { 0, 2, 3 }, "drop-final" => new[] { 0, 1, 2 },
            "reorder" => new[] { 3, 0, 2 }, _ => new[] { 0, 1, 1, 2, 3 }
        };
        var requested = order.Select(index => detached ? original.Paragraphs[index].WithTabStops(original.Paragraphs[index].TabStops)
            : original.Paragraphs[index]).ToArray();
        document.ReplaceStoryParagraphs(original.Id, requested);
        byte[] bytes = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument saved = PublisherDocument.Load(bytes);
        PublisherTextStory actual = saved.TextStories[0];
        Assert.Equal(string.Join("\n", requested.Select(p => string.Concat(p.Runs.Select(r => r.Text)))) + "\n", actual.Text);
        Assert.Equal(order.Length, actual.Paragraphs.Count);
        for (int i = 0; i < order.Length; i++) {
            OfficeRichTextParagraph expected = requested[i], p = actual.Paragraphs[i];
            Assert.Equal((expected.Alignment, expected.Margins, expected.Indent, expected.LineHeight, expected.LineHeightFactor),
                (p.Alignment, p.Margins, p.Indent, p.LineHeight, p.LineHeightFactor));
            Assert.Equal(expected.Runs.Select(r => (r.Text, r.FontFamily, r.FontSize, r.Color, r.Bold, r.Italic)),
                p.Runs.Select(r => (r.Text, r.FontFamily, r.FontSize, r.Color, r.Bold, r.Italic)));
            Assert.Equal(Definition(document, source, original.Id, order[i], "FDPP"), Definition(saved, bytes, original.Id, i, "FDPP"));
            Assert.Equal(Definition(document, source, original.Id, order[i], "FDPC"), Definition(saved, bytes, original.Id, i, "FDPC"));
        }
        Assert.Equal(document.TextStories.Skip(1).Select(s => (s.Id, s.Text)), saved.TextStories.Skip(1).Select(s => (s.Id, s.Text)));
        PublisherNativeWritingTests.AssertPreservedStorage(source, bytes, "Quill/QuillSub/CONTENTS");
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(0, true)]
    public void Native_paragraph_identity_disambiguates_equal_visible_values_with_different_language_definitions(int selected, bool detached) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(Fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory original = document.TextStories.Last();
        Assert.True(PublisherTextComparison.ParagraphEqual(original.Paragraphs[0], original.Paragraphs[1]));
        Assert.NotEqual(Definition(document, source, original.Id, 0, "FDPC"), Definition(document, source, original.Id, 1, "FDPC"));
        OfficeRichTextParagraph p = original.Paragraphs[selected];
        document.ReplaceStoryParagraphs(original.Id, new[] { detached ? p.WithTabStops(p.TabStops) : p, original.Paragraphs[2] });
        if (detached) {
            var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
            Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
        } else {
            byte[] bytes = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            PublisherDocument saved = PublisherDocument.Load(bytes);
            Assert.Equal(original.Paragraphs[selected].Runs[0].Text + "\n" + original.Paragraphs[2].Runs[0].Text + "\n", saved.TextStories.Last().Text);
            Assert.Equal(Definition(document, source, original.Id, selected, "FDPC"), Definition(saved, bytes, original.Id, 0, "FDPC"));
        }
    }

    private static byte[] Definition(PublisherDocument document, byte[] source, uint storyId, int paragraph, string tag) {
        // Source style offsets describe the retained input, independently of the edited model.
        PublisherQuillStoryStyles binding = document.SourceStyleCatalog.Stories.Single(s => s.StoryId == storyId);
        PublisherQuillStyleUse[] pp = binding.Uses.Where(u => u.Tag == "FDPP").ToArray();
        int start = pp[paragraph].NativeStart;
        PublisherQuillSourceSpan span = binding.Uses.Single(u => u.Tag == tag && u.NativeStart <= start
            && u.NativeStart + u.NativeLength > start).DirectDefinition;
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        byte[] quill = compound!.Streams["Quill/QuillSub/CONTENTS"];
        return quill.Skip(span.Offset).Take(span.Length).ToArray();
    }
}
