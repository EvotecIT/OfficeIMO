using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherParagraphFormatWritingTests {
    [Theory]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("justify")]
    [InlineData("spacing")]
    [InlineData("insets")]
    [InlineData("first")]
    [InlineData("hanging")]
    [InlineData("exact")]
    [InlineData("factor")]
    [InlineData("single")]
    [InlineData("tabs")]
    [InlineData("clear-tabs")]
    [InlineData("combined")]
    [InlineData("per-paragraph")]
    [InlineData("topology")]
    [InlineData("linked")]
    [InlineData("chained-reset")]
    [InlineData("allocation")]
    [InlineData("fractional-spacing")]
    public void Public_native_saves_retain_requested_paragraph_layout_and_unrelated_storage(string variant) {
        string fixture = variant is "single" or "clear-tabs" or "chained-reset" ? "TextParagraphAttributes/native-textparagraphattributes.pub"
            : variant == "linked" ? "Newsletter/native-newsletter.pub" : variant == "per-paragraph" ? "TextParagraphs/native-textparagraphs.pub" : "Text/native-text.pub";
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        uint id = variant == "single" ? 10U : variant == "clear-tabs" ? 14U : variant == "chained-reset" ? 15U : variant == "linked" ? 3U : 2U;
        PublisherTextStory story = document.TextStories.Single(item => item.Id == id);
        int count = variant == "topology" ? 4 : variant == "allocation" ? 240 : story.Paragraphs.Count;
        var requested = Enumerable.Range(0, count).Select(index => Format(story.Paragraphs[variant is "topology" or "allocation" ? 0 : index],
            variant == "per-paragraph" ? new[] { "center", "exact", "hanging" }[index] : variant, index)).ToArray();
        document.ReplaceStoryParagraphs(id, requested);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        var result = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.False(result.Report.HasErrors); Assert.True(result.Report.HasLoss);
        PublisherDocument saved = PublisherDocument.Load(result.Value);
        Assert.Equal(document.TextStories.Select(item => item.Text), saved.TextStories.Select(item => item.Text));
        var actual = saved.TextStories.Single(item => item.Id == id).Paragraphs;
        Assert.Equal(requested.Length, actual.Count);
        for (int index = 0; index < actual.Count; index++) {
            var a = requested[index]; var b = actual[index];
            Assert.Equal((a.Alignment, a.LineHeight, a.LineHeightFactor, a.Margins, a.Indent), (b.Alignment, b.LineHeight, b.LineHeightFactor, b.Margins, b.Indent));
            Assert.Equal(a.Runs.Select(run => (run.Text, run.FontFamily, run.FontSize, run.Color)), b.Runs.Select(run => (run.Text, run.FontFamily, run.FontSize, run.Color)));
            Assert.Equal((a.TabStops!.Origin, a.TabStops.DefaultInterval, a.TabStops.AlignWithParagraph),
                (b.TabStops!.Origin, b.TabStops.DefaultInterval, b.TabStops.AlignWithParagraph));
            Assert.Equal(a.TabStops.Stops.Select(stop => (stop.Position, stop.Alignment, stop.Character)), b.TabStops.Stops.Select(stop => (stop.Position, stop.Alignment, stop.Character)));
        }
        Assert.Equal(document.Pages.SelectMany(page => page.TextFrames).Select(frame => (frame.Id, frame.StoryId, frame.PageId, frame.PreviousFrameId, frame.NextFrameId, frame.X, frame.Y, frame.Width, frame.Height)),
            saved.Pages.SelectMany(page => page.TextFrames).Select(frame => (frame.Id, frame.StoryId, frame.PageId, frame.PreviousFrameId, frame.NextFrameId, frame.X, frame.Y, frame.Width, frame.Height)));
        PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Quill/QuillSub/CONTENTS");
        Assert.Equal(source, PublisherDocument.Load(source).ToPubBytes());
    }

    [Theory]
    [InlineData("tab-leader")]
    [InlineData("tab-origin")]
    [InlineData("two-indents")]
    [InlineData("non-native-unit")]
    public void Native_unrepresentable_paragraph_settings_refuse_output_even_when_loss_is_allowed(string variant) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"));
        var source = document.TextStories[0].Paragraphs[0];
        var stops = new OfficeTextTabStops(new[] { variant == "tab-leader" ? new OfficeTextTabStop(72).WithLeader(".") : new OfficeTextTabStop(72) },
            origin: variant == "tab-origin" ? 10 : 0);
        var paragraph = new OfficeRichTextParagraph(source.Runs, source.Alignment, source.LineHeight,
            variant == "non-native-unit" ? new OfficeTextPadding(0, 1D / 3, 0, 0) : source.Margins,
            variant == "two-indents" ? new OfficeTextParagraphIndent(12, 18) : source.Indent, source.LineHeightFactor).WithTabStops(stops);
        document.ReplaceStoryParagraphs(2, new[] { paragraph });
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal(paragraph.Margins, document.TextStories[0].Paragraphs[0].Margins);
    }

    internal static OfficeRichTextParagraph Format(OfficeRichTextParagraph source, string variant, int ordinal = 0) {
        OfficeTextAlignment alignment = variant is "center" or "combined" or "topology" or "linked" ? OfficeTextAlignment.Center
            : variant == "right" ? OfficeTextAlignment.Right : variant == "justify" ? OfficeTextAlignment.Justify : source.Alignment;
        var margins = variant is "spacing" or "combined" ? new OfficeTextPadding(variant == "combined" ? 36 : source.Margins.Left, 6, variant == "combined" ? 18 : source.Margins.Right, 9)
            : variant == "insets" ? new OfficeTextPadding(36, source.Margins.Top, 24, source.Margins.Bottom)
            : variant == "hanging" ? new OfficeTextPadding(18, source.Margins.Top, source.Margins.Right, source.Margins.Bottom)
            : variant == "chained-reset" ? OfficeTextPadding.Empty
            : variant == "allocation" ? new OfficeTextPadding(source.Margins.Left, ordinal + 1, source.Margins.Right, source.Margins.Bottom)
            : variant == "fractional-spacing" ? new OfficeTextPadding(source.Margins.Left, 1.21, source.Margins.Right, 2.42) : source.Margins;
        OfficeTextParagraphIndent indent = variant is "first" or "combined" ? OfficeTextParagraphIndent.FirstLine(18)
            : variant == "hanging" ? OfficeTextParagraphIndent.Hanging(18) : variant == "chained-reset" ? OfficeTextParagraphIndent.Empty : source.Indent;
        double? height = variant == "exact" ? 18 : variant is "factor" or "single" or "combined" or "linked" or "chained-reset" ? null : source.LineHeight;
        double? factor = variant is "factor" or "combined" or "linked" ? 1.5 : variant is "single" or "chained-reset" ? 1 : variant == "exact" ? null : source.LineHeightFactor;
        var tabs = new OfficeTextTabStops(variant is "tabs" or "combined" ? new[] { new OfficeTextTabStop(72), new OfficeTextTabStop(144) }
            : variant is "clear-tabs" or "chained-reset" ? Array.Empty<OfficeTextTabStop>() : source.TabStops!.Stops, origin: -margins.Left);
        if (variant == "chained-reset") alignment = OfficeTextAlignment.Left;
        return new OfficeRichTextParagraph(source.Runs, alignment, height, margins, indent, factor).WithTabStops(tabs);
    }
}
