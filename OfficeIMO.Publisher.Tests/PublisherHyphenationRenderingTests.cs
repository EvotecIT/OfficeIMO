using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Pdf;
using System.Xml.Linq;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherHyphenationRenderingTests {
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
