using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;
using System.Xml.Linq;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherHyphenationRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unspecified_language_retains_native_body_and_cell_scene_language(bool cells) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(cells
            ? "NativeComparison/TableCellTyping/native-tablecelltyping.pub"
            : "NativeComparison/FrameCreation/native-framecreation.pub"));
        PublisherTable? table = cells ? document.Pages.SelectMany(page => page.Tables).First(t => t.Cells.Any(c =>
            c.Paragraphs.SelectMany(p => p.Runs).Any(r => r.LanguageId == 1036))) : null;
        PublisherTableCell? cell = table?.Cells.First(c => c.Paragraphs.SelectMany(p => p.Runs).Any(r => r.LanguageId == 1036));
        uint storyId = cells ? table!.StoryId!.Value : document.TextStories[0].Id;
        IReadOnlyList<OfficeRichTextParagraph> paragraphs = cells ? cell!.Paragraphs
            : document.TextStories.Single(s => s.Id == storyId).Paragraphs;
        string text = string.Concat(paragraphs[0].Runs.Select(r => r.Text));
        uint expectedLanguage = cells ? 1036U : 1033U;
        var requested = paragraphs.Select(p => new OfficeRichTextParagraph(p.Runs.Select(r => r.WithLanguageId(null)).ToArray(),
            p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops)).ToArray();
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, cell!.RowIndex, cell.ColumnIndex, requested);
        else document.ReplaceStoryParagraphs(storyId, requested);
        Assert.All(requested.SelectMany(p => p.Runs), r => Assert.Null(r.LanguageId));
        foreach (PublisherDocument current in new[] { document, PublisherDocument.Load(document.ToPubBytes(
            new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow })) }) {
            OfficeRichTextParagraph painted = Assert.Single(current.Pages.Where(page => table == null || page.Id == table.PageId)
                .SelectMany(page => PublisherNativeTests.Elements(page.CreateDrawing()))
                .OfType<OfficeDrawingRichText>().SelectMany(frame => frame.Paragraphs), p => string.Concat(p.Runs.Select(r => r.Text)) == text);
            Assert.All(painted.Runs, r => Assert.Equal(expectedLanguage, r.LanguageId));
        }
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData(1036U, false)]
    public void Retained_or_explicit_language_controls_hyphenation_before_and_after_native_save(uint? language, bool hyphenates) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TextHyphenation/native-texthyphenation.pub"));
        uint id = Assert.Single(document.Pages[0].TextFrames).StoryId;
        PublisherTextStory story = document.TextStories.Single(s => s.Id == id);
        document.ReplaceStoryParagraphs(id, story.Paragraphs.Select(p => new OfficeRichTextParagraph(
            p.Runs.Select(run => run.WithLanguageId(language)).ToArray(), p.Alignment, p.LineHeight, p.Margins,
            p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops)).ToArray());
        Assert.All(document.TextStories.Single(s => s.Id == id).Paragraphs.SelectMany(p => p.Runs),
            run => Assert.Equal(language, run.LanguageId));
        Assert.Equal(hyphenates, DisplayText(document).Contains("cre-"));
        byte[] saved = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument reloaded = PublisherDocument.Load(saved);
        Assert.Equal(hyphenates, DisplayText(reloaded).Contains("cre-"));
        Assert.Equal(story.Text, reloaded.TextStories.Single(s => s.Id == id).Text);
        Assert.All(reloaded.TextStories.Single(s => s.Id == id).Paragraphs.SelectMany(p => p.Runs),
            run => Assert.Equal(language ?? 1033U, run.LanguageId));
    }

    [Fact]
    public void Linked_frame_suffix_paints_without_consuming_the_next_source_character() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TextHyphenation/native-linked-hyphen-boundary.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextFrame ending = Assert.Single(document.Pages[14].TextFrames);
        PublisherTextFrame continuation = Assert.Single(document.Pages[13].TextFrames);
        PublisherTextStory story = document.TextStories.Single(s => s.Id == ending.StoryId);
        Assert.Equal(continuation.Id, ending.NextFrameId);
        Assert.Equal(ending.TextStart + ending.TextLength, continuation.TextStart);
        string assigned = story.Text.Substring(ending.TextStart!.Value, ending.TextLength!.Value);
        Assert.DoesNotContain("-", assigned);
        Assert.Equal(assigned + "-", DisplayText(document, 14));
        Assert.Equal(assigned + "-", PdfReadDocument.Open(document.ToPdfBytes()).Pages[14].ExtractText().Trim());
        Assert.Equal(source, document.ToPubBytes());
        Assert.DoesNotContain("-", story.Text);
    }

    [Fact]
    public void Native_hyphenation_settings_change_display_without_inserting_source_characters() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TextHyphenation/native-texthyphenation.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        uint id = Assert.Single(document.Pages[0].TextFrames).StoryId;
        PublisherTextStory story = document.TextStories.Single(s => s.Id == id);
        Assert.DoesNotContain("-", story.Text);
        Assert.Contains("cre-", DisplayText(document));
        Assert.Contains("cre-", PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].ExtractText());
        Assert.Equal(source, document.ToPubBytes());
        Assert.Equal(story.Text, document.TextStories.Single(s => s.Id == id).Text);

        document.SetStoryAutomaticHyphenation(id, false);
        Assert.DoesNotContain("-", DisplayText(document));
        Assert.DoesNotContain("-", PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].ExtractText());
        document.SetStoryAutomaticHyphenation(id, true);
        document.ReplaceStoryParagraphs(id, story.Paragraphs.Select(p => p.WithAutomaticHyphenationSuppressed(true)).ToArray());
        Assert.DoesNotContain("-", DisplayText(document));
        Assert.DoesNotContain("-", PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].ExtractText());
        Assert.Equal(story.Text, document.TextStories.Single(s => s.Id == id).Text);
    }

    private static string DisplayText(PublisherDocument document, int page = 0) => string.Concat(XDocument.Parse(document.ToSvg(page)).Descendants()
        .Where(e => e.Name.LocalName == "text" || e.Name.LocalName == "g" && e.Attribute("role")?.Value == "img")
        .Select(e => e.Name.LocalName == "text" ? e.Value : e.Attribute("aria-label")?.Value));
}
