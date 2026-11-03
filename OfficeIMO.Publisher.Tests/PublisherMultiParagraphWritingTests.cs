using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherMultiParagraphWritingTests {
    [Fact]
    public void Native_edited_fresh_multi_paragraph_story_supports_public_restoration_with_chain_and_storage_preserved() {
        byte[] source = Fixture("GeneratedMultiParagraph/native-edited.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory story = Assert.Single(document.TextStories);
        Assert.Equal(1359, story.Text.Length);
        Assert.StartsWith("Nulti-paragraph heading. ", story.Text);
        Assert.Equal(3, story.Paragraphs.Count);
        OfficeRichTextRun empty = Assert.Single(story.Paragraphs[1].Runs);
        Assert.Equal((string.Empty, 14D, true), (empty.Text, empty.FontSize, empty.Bold));
        Assert.Equal(source, document.ToPubBytes());
        Compare(PublisherDocument.Load(Fixture("GeneratedMultiParagraph/native-control-edited.pub")), document);

        OfficeRichTextParagraph[] paragraphs = story.Paragraphs.ToArray();
        OfficeRichTextParagraph first = paragraphs[0];
        OfficeRichTextRun[] runs = first.Runs.ToArray();
        OfficeRichTextRun run = runs[0];
        Assert.StartsWith("N", run.Text);
        runs[0] = new OfficeRichTextRun("M" + run.Text.Substring(1), run.FontSize, run.Color, run.Bold,
            run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle,
            run.StrikethroughStyle, run.Baseline, run.ParagraphIndent).WithLanguageId(run.LanguageId)
            .WithAutomaticPairKerningThreshold(run.AutomaticPairKerningThreshold);
        OfficeRichTextParagraph replacement = new OfficeRichTextParagraph(runs, first.Alignment, first.LineHeight, first.Margins,
            first.Indent, first.LineHeightFactor).WithTabStops(first.TabStops);
        paragraphs[0] = first.SuppressAutomaticHyphenation is bool suppressed
            ? replacement.WithAutomaticHyphenationSuppressed(suppressed) : replacement;
        document.ReplaceStoryParagraphs(story.Id, paragraphs);

        byte[] saved = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument restored = PublisherDocument.Load(saved);
        Assert.Equal(first.Runs.Select(item => item.Text.Length),
            Assert.Single(restored.TextStories).Paragraphs[0].Runs.Select(item => item.Text.Length));
        Compare(PublisherDocument.Load(Fixture("MultiParagraphCreationLinked/native-multiparagraphcreationlinked.pub")), restored);
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved, "Quill/QuillSub/CONTENTS");
        Assert.Equal(saved, restored.ToPubBytes());
    }

    private static void Compare(PublisherDocument expected, PublisherDocument actual) {
        PublisherTextStory before = Assert.Single(expected.TextStories), after = Assert.Single(actual.TextStories);
        Assert.Equal((before.Text, before.AutomaticHyphenation), (after.Text, after.AutomaticHyphenation));
        Assert.Equal(before.Paragraphs.Count, after.Paragraphs.Count);
        for (int index = 0; index < before.Paragraphs.Count; index++) {
            OfficeRichTextParagraph a = before.Paragraphs[index], b = after.Paragraphs[index];
            Assert.Equal((a.Alignment, a.LineHeight, a.LineHeightFactor, a.Margins, a.Indent, a.SuppressAutomaticHyphenation),
                (b.Alignment, b.LineHeight, b.LineHeightFactor, b.Margins, b.Indent, b.SuppressAutomaticHyphenation));
            Assert.Null(a.Label); Assert.Null(b.Label);
            OfficeTextTabStops nativeTabs = Assert.IsType<OfficeTextTabStops>(a.TabStops);
            OfficeTextTabStops generatedTabs = Assert.IsType<OfficeTextTabStops>(b.TabStops);
            // These tab-free producer profiles retain different repeating defaults.
            Assert.Equal(28.3D, nativeTabs.DefaultInterval, 10);
            Assert.Equal(36D, generatedTabs.DefaultInterval);
            Assert.Equal((nativeTabs.Origin, nativeTabs.AlignWithParagraph), (generatedTabs.Origin, generatedTabs.AlignWithParagraph));
            Assert.Empty(nativeTabs.Stops); Assert.Empty(generatedTabs.Stops);
            Assert.Equal(CharacterStyles(a), CharacterStyles(b));
        }
        CompareChain(expected, actual);
    }

    private static IEnumerable<object> CharacterStyles(OfficeRichTextParagraph paragraph) {
        foreach (OfficeRichTextRun run in paragraph.Runs) {
            if (run.Text.Length == 0) yield return CharacterStyle(null, run);
            foreach (char character in run.Text) yield return CharacterStyle(character, run);
        }
        // Empty-run sentinels retain typing styles; the final newline retains paragraph-terminator styling.
        yield return CharacterStyle('\n', paragraph.Runs.Last());
    }

    private static object CharacterStyle(char? character, OfficeRichTextRun run) => new {
        Character = character, run.FontSize, run.Color, run.FontFamily, run.Bold, run.Italic,
        run.UnderlineStyle, run.StrikethroughStyle, run.Baseline, run.BackgroundColor,
        run.ParagraphIndent, run.LanguageId, run.AutomaticPairKerningThreshold
    };

    private static void CompareChain(PublisherDocument expected, PublisherDocument actual) {
        PublisherTextFrame[] before = Chain(expected), after = Chain(actual);
        var expectedPages = expected.Pages.Select((page, index) => (page.Id, index)).ToDictionary(item => item.Id, item => item.index);
        var actualPages = actual.Pages.Select((page, index) => (page.Id, index)).ToDictionary(item => item.Id, item => item.index);
        Assert.Equal(expected.Pages.Select(page => (page.Width, page.Height)), actual.Pages.Select(page => (page.Width, page.Height)));
        for (int index = 0; index < before.Length; index++) {
            PublisherTextFrame a = before[index], b = after[index];
            Assert.Equal((expectedPages[a.PageId], a.Order), (actualPages[b.PageId], b.Order));
            Assert.Equal((a.X, a.Y, a.Width, a.Height, a.PageTransform), (b.X, b.Y, b.Width, b.Height, b.PageTransform));
            Assert.Equal((a.ColumnCount, a.ColumnSpacing, a.Padding, a.VerticalAlignment),
                (b.ColumnCount, b.ColumnSpacing, b.Padding, b.VerticalAlignment));
            Assert.Equal((a.TextStart, a.TextLength, a.HasOverflow, a.CanEditLayout, a.LayoutEditLimitation),
                (b.TextStart, b.TextLength, b.HasOverflow, b.CanEditLayout, b.LayoutEditLimitation));
        }
    }

    private static PublisherTextFrame[] Chain(PublisherDocument document) {
        Assert.Equal(3, document.Pages.Count);
        PublisherTextStory story = Assert.Single(document.TextStories);
        PublisherTextFrame[] frames = document.Pages.Select(page => {
            PublisherTextFrame frame = Assert.Single(page.TextFrames);
            Assert.Equal(page.Id, frame.PageId);
            return frame;
        }).OrderBy(frame => frame.Order).ToArray();
        Assert.Equal(new uint[] { 0, 1, 2 }, frames.Select(frame => frame.Order));
        Assert.Equal(new[] { document.Pages[0].Id, document.Pages[2].Id, document.Pages[1].Id }, frames.Select(frame => frame.PageId));
        for (int index = 0; index < frames.Length; index++) {
            PublisherTextFrame frame = frames[index];
            Assert.Equal(story.Id, frame.StoryId);
            Assert.Equal(index == 0 ? null : (uint?)frames[index - 1].Id, frame.PreviousFrameId);
            Assert.Equal(index == frames.Length - 1 ? null : (uint?)frames[index + 1].Id, frame.NextFrameId);
            Assert.Empty(frame.WrapObjectIds);
        }
        return frames;
    }

    private static byte[] Fixture(string path) => File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + path));
}
