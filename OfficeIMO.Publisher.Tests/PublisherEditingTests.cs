using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Internal;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherEditingTests {
    [Fact]
    public void Story_and_bounds_edits_reach_svg_pdf_and_current_snapshots() {
        PublisherDocument document = Load("Text");
        PublisherTextStory oldStory = Assert.Single(document.TextStories);
        PublisherTextFrame oldFrame = Assert.Single(document.Pages[0].TextFrames);
        var edited = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("Edited publication", 18, OfficeColor.FromRgb(180, 20, 30), bold: true, fontFamily: "Arial")
        });
        document.ReplaceStoryParagraphs(oldStory.Id, new[] { edited });
        document.SetTextFrameBounds(oldFrame.Id, 90, 108, 300, 96);

        PublisherTextStory story = Assert.Single(document.TextStories);
        PublisherTextFrame frame = Assert.Single(document.Pages[0].TextFrames);
        Assert.Equal(oldStory.Id, story.Id); Assert.Equal(oldFrame.Id, frame.Id); Assert.Equal(oldFrame.StoryId, frame.StoryId);
        Assert.Equal("Edited publication\n", story.Text);
        Assert.Equal("Publisher text\n", oldStory.Text);
        Assert.Equal(72, oldFrame.X);
        Assert.Equal(90, frame.X, 6); Assert.Equal(108, frame.Y, 6); Assert.Equal(300, frame.Width, 6); Assert.Equal(96, frame.Height, 6);
        OfficeDrawingRichText paint = Assert.Single(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingRichText>());
        OfficeRichTextRun run = Assert.Single(Assert.Single(paint.Paragraphs).Runs);
        Assert.True(run.Bold); Assert.Equal(18, run.FontSize); Assert.Equal(OfficeColor.FromRgb(180, 20, 30), run.Color);
        Assert.Contains("Edited publication", document.ToSvg());
        Assert.Contains("Edited publication", PdfReadDocument.Open(document.ToPdfBytes()).ExtractText());

        document.ReplaceStoryParagraphs(story.Id, Paragraphs("Second version"));
        Assert.Contains("Second version", document.ToSvg());
        Assert.DoesNotContain("Edited publication", document.ToSvg());
        Assert.Contains("Second version", PdfReadDocument.Open(document.ToPdfBytes()).ExtractText());
    }

    [Theory]
    [InlineData("NativeComparison/Newsletter/native-newsletter.pub", 3U)]
    [InlineData("SampleNewsletter.pub", 18U)]
    public void Long_then_short_edits_keep_chain_identity_and_recalculate_overflow(string file, uint storyId) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(file));
        PublisherTextFrame[] before = Frames(document, storyId);
        Assert.True(before.Length >= 2);
        foreach (PublisherTextFrame frame in before)
            document.SetTextFrameBounds(frame.Id, frame.X, frame.Y, 90, 32);
        document.ReplaceStoryParagraphs(storyId, Paragraphs(string.Concat(Enumerable.Repeat("Continuation remains complete. ", 80))));
        PublisherTextFrame[] longer = Frames(document, storyId);
        Assert.Equal(before.Select(frame => (frame.Id, frame.StoryId, frame.PageId, frame.PreviousFrameId, frame.NextFrameId)),
            longer.Select(frame => (frame.Id, frame.StoryId, frame.PageId, frame.PreviousFrameId, frame.NextFrameId)));
        Assert.True(longer[longer.Length - 1].HasOverflow);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_FRAME_OVERFLOW");
        AssertRanges(document, storyId);

        document.ReplaceStoryParagraphs(storyId, Paragraphs("Short"));
        PublisherTextFrame[] shorter = Frames(document, storyId);
        Assert.False(shorter[shorter.Length - 1].HasOverflow);
        AssertRanges(document, storyId);
        Assert.All(shorter.Skip(1), frame => Assert.Equal(0, frame.TextLength));
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_FRAME_OVERFLOW"
            && shorter.Any(frame => item.Location == "Contents/object/" + frame.Id));
    }

    [Fact]
    public void Paragraph_normalization_and_empty_story_keep_one_text_contract() {
        PublisherDocument document = Load("Text");
        uint id = document.TextStories[0].Id;
        document.ReplaceStoryParagraphs(id, Paragraphs("First\r\nline\rlast"));
        PublisherTextStory story = document.TextStories[0];
        Assert.Equal("First\nline\nlast\n", story.Text);
        Assert.Equal("First\nline\nlast", story.Paragraphs[0].Runs[0].Text);
        document.ReplaceStoryParagraphs(id, Array.Empty<OfficeRichTextParagraph>());
        Assert.Empty(document.TextStories[0].Text); Assert.Empty(document.TextStories[0].Paragraphs);
        Assert.All(document.Pages[0].TextFrames, frame => { Assert.Equal(0, frame.TextStart); Assert.Equal(0, frame.TextLength); Assert.False(frame.HasOverflow); });
        document.ReplaceStoryParagraphs(id, Paragraphs("Restored"));
        Assert.Equal("Restored\n", document.TextStories[0].Text);
        Assert.Contains("Restored", document.ToSvg());
    }

    [Theory]
    [InlineData(double.NaN, 20, 30, 40)]
    [InlineData(10, double.PositiveInfinity, 30, 40)]
    [InlineData(10, 20, 0, 40)]
    [InlineData(10, 20, 30, -1)]
    [InlineData(double.MaxValue, 20, 30, 40)]
    public void Invalid_bounds_leave_the_publication_unchanged(double x, double y, double width, double height) {
        PublisherDocument document = Load("Text");
        var pages = document.Pages; var report = document.ReadReport; string before = document.ToSvg();
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetTextFrameBounds(pages[0].TextFrames[0].Id, x, y, width, height));
        Assert.Same(pages, document.Pages); Assert.Same(report, document.ReadReport); Assert.Equal(before, document.ToSvg());
    }

    [Fact]
    public void Failed_or_cancelled_story_edits_do_not_replace_current_state() {
        PublisherDocument document = Load("Text");
        uint storyId = document.TextStories[0].Id;
        var stories = document.TextStories; var pages = document.Pages;
        Assert.Throws<ArgumentException>(() => document.ReplaceStoryParagraphs(storyId, Paragraphs("\uD800")));
        Assert.Throws<ArgumentException>(() => document.ReplaceStoryParagraphs(storyId, Paragraphs(new string('x', 100001))));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.ReplaceStoryParagraphs(uint.MaxValue, Paragraphs("missing")));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.ReplaceStoryParagraphs(storyId, Paragraphs("cancelled"), cancelled.Token));
        Assert.Throws<OperationCanceledException>(() => document.SetTextFrameBounds(pages[0].TextFrames[0].Id, 10, 20, 30, 40, cancelled.Token));
        Assert.Same(stories, document.TextStories); Assert.Same(pages, document.Pages);
    }

    [Fact]
    public void Table_grouped_and_unresolved_objects_are_explicitly_rejected() {
        PublisherDocument document = Load("TableLabel");
        PublisherTable table = Assert.Single(document.Pages[0].Tables);
        var stories = document.TextStories;
        Assert.Throws<NotSupportedException>(() => document.ReplaceStoryParagraphs(table.StoryId!.Value, Paragraphs("unsafe cell edit")));
        Assert.Throws<NotSupportedException>(() => document.SetTextFrameBounds(table.Id, 10, 20, 30, 40));
        Assert.Throws<NotSupportedException>(() => document.SetTextFrameBounds(uint.MaxValue, 10, 20, 30, 40));
        Assert.Same(stories, document.TextStories);
        PublisherDocument grouped = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        var groupedPages = grouped.Pages;
        Assert.Contains(groupedPages.SelectMany(page => page.TextFrames), frame => frame.Id == 345);
        Assert.Throws<NotSupportedException>(() => grouped.SetTextFrameBounds(345, 10, 20, 30, 40));
        Assert.Same(groupedPages, grouped.Pages);
    }

    [Fact]
    public void Drawing_copies_and_caller_input_cannot_change_the_model() {
        byte[] bytes = File.ReadAllBytes(PublisherNativeTests.Fixture(Path.Combine("NativeComparison", "Text", "native-text.pub")));
        using var loadCancellation = new CancellationTokenSource();
        PublisherDocument document = PublisherDocument.Load(bytes, cancellationToken: loadCancellation.Token);
        loadCancellation.Cancel(); Array.Clear(bytes, 0, bytes.Length);
        string before = document.ToSvg();
        OfficeDrawing copy = document.Pages[0].CreateDrawing();
        copy.AddShape(new OfficeShape { Kind = OfficeShapeKind.Rectangle, Width = 12, Height = 12, FillColor = OfficeColor.Black }, 0, 0);
        Assert.Equal(before, document.ToSvg());
        Assert.NotEqual(before, OfficeDrawingSvgExporter.ToSvg(copy));
        document.ReplaceStoryParagraphs(document.TextStories[0].Id, Paragraphs("Independent"));
        Assert.Contains("Independent", document.ToSvg());
    }

    [Theory]
    [InlineData("invalid-utf16")]
    [InlineData("null-character")]
    [InlineData("control-marker")]
    [InlineData("nonpositive-font-size")]
    [InlineData("nonfinite-font-size")]
    public void Edited_labels_reject_invalid_text_and_font_sizes_atomically(string scenario) {
        PublisherDocument document = Load("Text");
        string marker = scenario switch { "invalid-utf16" => "\uD800", "null-character" => "\0", "control-marker" => "\u0001", _ => "•" };
        double size = scenario switch { "nonpositive-font-size" => 0, "nonfinite-font-size" => double.NaN, _ => 12 };
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun(marker, size, OfficeColor.Black), 0);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Body", 12, OfficeColor.Black) }, label);
        var stories = document.TextStories; var pages = document.Pages;
        Assert.Throws<ArgumentException>(() => document.ReplaceStoryParagraphs(stories[0].Id, new[] { paragraph }));
        Assert.Same(stories, document.TextStories); Assert.Same(pages, document.Pages);
    }

    [Theory]
    [InlineData("characters")]
    [InlineData("runs")]
    public void Edited_labels_obey_story_character_and_run_limits(string budget) {
        PublisherDocument document = Load("Text");
        int count = budget == "characters" ? 25 : 1366;
        string marker = budget == "characters" ? new string('x', 4000) : "•";
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun(marker, 12, OfficeColor.Black), 0);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("b", 12, OfficeColor.Black) }, label);
        var stories = document.TextStories; var pages = document.Pages;
        Assert.Throws<ArgumentException>(() => document.ReplaceStoryParagraphs(stories[0].Id, Enumerable.Repeat(paragraph, count).ToArray()));
        Assert.Same(stories, document.TextStories); Assert.Same(pages, document.Pages);
    }

    [Fact]
    public void Edited_labels_count_toward_the_publication_character_limit() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Newsletter/native-newsletter.pub"),
            new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 12000 } });
        Assert.True(document.TextStories.Count >= 2);
        uint[] ids = document.TextStories.Take(2).Select(story => story.Id).ToArray();
        // Unplaced labels still belong to the publication's retained content;
        // projection accounting cannot replace this source admission check.
        foreach (PublisherTextFrame frame in document.Pages.SelectMany(page => page.TextFrames).ToArray())
            document.SetTextFrameBounds(frame.Id, frame.X, frame.Y, 1, 1);
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun(new string('x', 4000), 12, OfficeColor.Black), 0);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Body", 12, OfficeColor.Black) }, label);
        var replacement = new[] { paragraph, paragraph };
        document.ReplaceStoryParagraphs(ids[0], replacement);
        var stories = document.TextStories; var pages = document.Pages;
        Assert.Throws<ArgumentException>(() => document.ReplaceStoryParagraphs(ids[1], replacement));
        Assert.Same(stories, document.TextStories); Assert.Same(pages, document.Pages);
    }

    [Fact]
    public void Projection_budget_failure_from_labels_leaves_previous_state_usable() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"),
            new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 200 } });
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun(new string('x', 100), 1, OfficeColor.Black), 0);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Body", 12, OfficeColor.Black) }, label);
        var stories = document.TextStories; var pages = document.Pages; string svg = document.ToSvg();
        Assert.Throws<InvalidDataException>(() => document.ReplaceStoryParagraphs(stories[0].Id, new[] { paragraph }));
        Assert.Same(stories, document.TextStories); Assert.Same(pages, document.Pages);
        Assert.Equal(svg, document.ToSvg());
    }

    private static PublisherDocument Load(string name) => PublisherDocument.Load(PublisherNativeTests.Fixture(Path.Combine("NativeComparison", name, "native-" + name.ToLowerInvariant() + ".pub")));
    private static OfficeRichTextParagraph[] Paragraphs(string text) => new[] { new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, 12, OfficeColor.Black, fontFamily: "Arial") }) };
    private static PublisherTextFrame[] Frames(PublisherDocument document, uint story) => document.Pages.SelectMany(page => page.TextFrames).Where(frame => frame.StoryId == story).OrderBy(frame => frame.Order).ToArray();
    private static void AssertRanges(PublisherDocument document, uint id) {
        PublisherTextFrame[] frames = Frames(document, id);
        Assert.Equal(0, frames[0].TextStart);
        for (int i = 1; i < frames.Length; i++) Assert.Equal(frames[i - 1].TextStart + frames[i - 1].TextLength, frames[i].TextStart);
        PublisherTextStory story = document.TextStories.Single(item => item.Id == id);
        string placed = string.Concat(frames.Select(frame => story.Text.Substring(frame.TextStart!.Value, frame.TextLength!.Value)));
        Assert.StartsWith(placed, story.Text);
        Assert.Equal(placed.Length < story.Text.Length, frames[frames.Length - 1].HasOverflow);
    }
}
