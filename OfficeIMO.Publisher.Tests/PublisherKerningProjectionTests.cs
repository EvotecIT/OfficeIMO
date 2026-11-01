using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherKerningProjectionTests {
    [Theory]
    [InlineData(14)]
    [InlineData(0)]
    public void Native_default_is_applied_only_to_scene_runs_and_preserves_public_metadata(int threshold) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(
            "NativeComparison/TextPairKerning/native-threshold-" + threshold + ".pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.All(document.TextStories.SelectMany(story => story.Paragraphs).SelectMany(paragraph => paragraph.Runs),
            run => Assert.Equal(threshold == 14 ? (double?)null : 0D, run.AutomaticPairKerningThreshold));
        OfficeRichTextRun[] painted = document.Pages.SelectMany(page => PublisherNativeTests.Elements(page.CreateDrawing()))
            .OfType<OfficeDrawingRichText>().SelectMany(frame => frame.Paragraphs).SelectMany(paragraph => paragraph.Runs).ToArray();
        Assert.NotEmpty(painted);
        Assert.All(painted, run => Assert.Equal((double)threshold, run.AutomaticPairKerningThreshold));
        Assert.Equal(source, document.ToPubBytes());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Unspecified_replacement_keeps_source_thresholds_in_body_and_cell_scenes(bool cells, bool merge) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(cells
            ? "NativeComparison/TableLabel/native-tablelabel.pub" : "NativeComparison/Text/native-text.pub"));
        PublisherTable? table = cells ? document.Pages[0].Tables.Single() : null;
        uint storyId = cells ? table!.StoryId!.Value : document.TextStories[0].Id;
        OfficeRichTextParagraph source = cells ? table!.Cells[0].Paragraphs[0]
            : document.TextStories.Single(story => story.Id == storyId).Paragraphs[0];
        string text = string.Concat(source.Runs.Select(run => run.Text));
        Assert.True(text.Length >= 2);
        var runs = merge ? new[] { Run(source.Runs[0], text.Substring(0, 2), 8), Run(source.Runs[0], text.Substring(2), 18) }
            : new[] { Run(source.Runs[0], text, 8) };
        Replace(document, table, storyId, Paragraph(source, runs));
        var options = new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow };
        document = PublisherDocument.Load(document.ToPubBytes(options));
        table = cells ? document.Pages[0].Tables.Single() : null;
        OfficeRichTextParagraph retained = cells ? table!.Cells[0].Paragraphs[0]
            : document.TextStories.Single(story => story.Id == storyId).Paragraphs[0];
        var requested = merge ? new[] { Run(retained.Runs[0], text, null) }
            : new[] { Run(retained.Runs[0], text + " extra", null) };
        string expectedText = requested[0].Text;
        Replace(document, table, storyId, Paragraph(retained, requested));
        OfficeRichTextParagraph current = cells ? document.Pages[0].Tables.Single().Cells[0].Paragraphs[0]
            : document.TextStories.Single(story => story.Id == storyId).Paragraphs[0];
        Assert.Single(current.Runs);
        Assert.Null(current.Runs[0].AutomaticPairKerningThreshold);
        Assert.Equal(expectedText, current.Runs[0].Text);
        double?[] expected = merge ? new double?[] { 8, 18 } : new double?[] { 8 };
        Assert.Equal(expected, Painted(document, expectedText).Select(run => run.AutomaticPairKerningThreshold));
        PublisherDocument saved = PublisherDocument.Load(document.ToPubBytes(options));
        Assert.Equal(expected, Painted(saved, expectedText).Select(run => run.AutomaticPairKerningThreshold));
        var recovered = cells ? saved.Pages[0].Tables.Single().Cells[0].Paragraphs[0]
            : saved.TextStories.Single(story => story.Id == storyId).Paragraphs[0];
        Assert.Equal(expected, recovered.Runs.Select(run => run.AutomaticPairKerningThreshold));
    }

    private static OfficeRichTextRun[] Painted(PublisherDocument document, string text) {
        OfficeRichTextParagraph paragraph = Assert.Single(document.Pages.SelectMany(page => PublisherNativeTests.Elements(page.CreateDrawing()))
            .OfType<OfficeDrawingRichText>().SelectMany(frame => frame.Paragraphs),
            paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)) == text);
        return paragraph.Runs.ToArray();
    }

    private static void Replace(PublisherDocument document, PublisherTable? table, uint id, OfficeRichTextParagraph paragraph) {
        if (table != null) document.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] { paragraph });
        else document.ReplaceStoryParagraphs(id, new[] { paragraph });
    }

    private static OfficeRichTextParagraph Paragraph(OfficeRichTextParagraph source, OfficeRichTextRun[] runs) =>
        new OfficeRichTextParagraph(runs, source.Alignment, source.LineHeight, source.Margins, source.Indent, source.LineHeightFactor)
            .WithTabStops(source.TabStops);

    private static OfficeRichTextRun Run(OfficeRichTextRun source, string text, double? threshold) =>
        new OfficeRichTextRun(text, source.FontSize, source.Color, source.Bold, source.Italic, source.Underline,
            source.FontFamily, source.Strikethrough, source.BackgroundColor, source.UnderlineStyle,
            source.StrikethroughStyle, source.Baseline, source.ParagraphIndent)
            .WithLanguageId(source.LanguageId).WithAutomaticPairKerningThreshold(threshold);
}
