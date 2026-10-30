using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherParagraphAttributeTests {
    [Fact]
    public void Native_tab_lines_follow_paragraph_center_and_right_alignment() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TextAlignedTabs/native-textalignedtabs.pub"));
        Assert.Single(document.Pages); Assert.Equal(3, document.TextStories.Count);
        var alignments = new[] { OfficeTextAlignment.Left, OfficeTextAlignment.Center, OfficeTextAlignment.Right };
        for (int index = 0; index < alignments.Length; index++) {
            PublisherTextStory story = document.TextStories[index];
            Assert.EndsWith("Left\tMiddle\tRight\n", story.Text);
            Assert.All(story.Paragraphs, paragraph => {
                Assert.Equal(alignments[index], paragraph.Alignment);
                Assert.Equal(index != 0, paragraph.TabStops!.AlignWithParagraph);
                Assert.Equal(new[] { 72D, 144D }, paragraph.TabStops.Stops.Select(stop => stop.Position));
            });
        }
    }

    [Fact]
    public void Native_paragraph_controls_retain_spacing_indents_alignment_and_multiple_indexed_tabs() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TextParagraphAttributes/native-textparagraphattributes.pub"));
        Assert.Equal(4, document.Pages.Count); Assert.Equal(14, document.TextStories.Count);
        string[] names = { "Default", "Center", "Right", "Justify", "Spacing", "Insets", "FirstLine", "Hanging",
            "Exactly18", "Multiple1pt5", "Single", "Double", "Tabs", "Combined" };
        for (int index = 0; index < names.Length; index++) {
            string name = names[index]; PublisherTextStory story = document.TextStories[index];
            Assert.StartsWith("Native paragraph control: " + name, story.Text); Assert.Equal(2, story.Paragraphs.Count);
            Assert.All(story.Paragraphs, paragraph => {
                Assert.Equal(name is "Center" or "Combined" ? OfficeTextAlignment.Center : name == "Right" ? OfficeTextAlignment.Right
                    : name == "Justify" ? OfficeTextAlignment.Justify : OfficeTextAlignment.Left, paragraph.Alignment);
                Assert.Equal(name is "Insets" or "Combined" ? 36 : name == "Hanging" ? 18 : 0, paragraph.Margins.Left);
                Assert.Equal(name is "Spacing" or "Combined" ? 6 : 0, paragraph.Margins.Top);
                Assert.Equal(name == "Insets" ? 24 : name == "Combined" ? 18 : 0, paragraph.Margins.Right);
                Assert.Equal(name is "Spacing" or "Combined" ? 9 : 0, paragraph.Margins.Bottom);
                Assert.Equal(name == "FirstLine" ? 18 : name == "Combined" ? 12 : 0, paragraph.Indent.FirstLineOffset);
                Assert.Equal(name == "Hanging" ? 18 : 0, paragraph.Indent.ContinuationLineOffset);
                Assert.Equal(name == "Exactly18" ? 18D : (double?)null, paragraph.LineHeight);
                Assert.Equal(name == "Exactly18" ? (double?)null : name == "Multiple1pt5" ? 1.5 : name == "Double" ? 2 : name == "Combined" ? 1.25 : 1,
                    paragraph.LineHeightFactor);
                Assert.NotNull(paragraph.TabStops); Assert.Equal(-paragraph.Margins.Left, paragraph.TabStops!.Origin);
                Assert.Equal(name is "Tabs" or "Combined" ? new[] { 72D, 144D } : Array.Empty<double>(), paragraph.TabStops.Stops.Select(stop => stop.Position));
                Assert.All(paragraph.TabStops.Stops, stop => Assert.Equal(OfficeTextTabAlignment.Left, stop.Alignment));
                Assert.Equal(36, paragraph.TabStops.DefaultInterval);
            });
        }
    }
}
