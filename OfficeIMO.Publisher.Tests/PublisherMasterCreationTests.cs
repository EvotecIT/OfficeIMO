using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherMasterCreationTests {
    [Theory]
    [InlineData("MultipleMasters/native-multiplemasters.pub", false)]
    [InlineData("TwoPageMaster/native-twopagemaster.pub", true)]
    [InlineData("GeneratedMultipleMasters/generated-multiplemasters.pub", false)]
    [InlineData("GeneratedMultipleMasters/native-normalized.pub", false)]
    [InlineData("GeneratedTwoPageMaster/generated-twopagemaster.pub", true)]
    [InlineData("GeneratedTwoPageMaster/native-normalized.pub", true)]
    public void Master_definitions_and_direct_assignments_survive_repeated_text_and_bounds_source_saves(string fixture, bool paired) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherMasterPageSide?[] sides = paired
            ? new PublisherMasterPageSide?[] { PublisherMasterPageSide.Single, PublisherMasterPageSide.Left, PublisherMasterPageSide.Right }
            : new PublisherMasterPageSide?[] { PublisherMasterPageSide.Single, PublisherMasterPageSide.Single };
        uint[] masters = document.MasterPages.Select(page => page.Id).ToArray();
        uint?[] assignments = { masters[0], masters[1], masters[paired ? 2 : 0] };
        var original = document.MasterPages.Concat(document.Pages).SelectMany(page => page.TextFrames).ToDictionary(frame => frame.Id);
        var masterFrames = document.MasterPages.SelectMany(page => page.TextFrames).Select(frame => frame.Id).ToHashSet();
        Assert.Equal(sides, document.MasterPages.Select(page => page.MasterSide));
        Assert.Equal(assignments, document.Pages.Select(page => page.MasterPageId));
        Assert.All(document.Pages, page => Assert.Single(page.TextFrames));
        Assert.All(document.MasterPages, page => Assert.Single(page.TextFrames));

        for (int iteration = 0; iteration < 2; iteration++) {
            var oldMasters = document.MasterPages;
            foreach (PublisherTextFrame frame in document.MasterPages.Concat(document.Pages).SelectMany(page => page.TextFrames).ToArray()) {
                document.SetTextFrameBounds(frame.Id, frame.X + 6, frame.Y + 12, frame.Width, frame.Height);
                if (!masterFrames.Contains(frame.Id)) continue;
                OfficeRichTextParagraph paragraph = document.TextStories.Single(story => story.Id == frame.StoryId).Paragraphs.Single();
                OfficeRichTextRun run = paragraph.Runs.Single();
                document.ReplaceStoryParagraphs(frame.StoryId, new[] {
                    new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(run.Text + (iteration == 0 ? " updated" : " restored"),
                        run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough,
                        run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle, run.Baseline, run.ParagraphIndent).WithLanguageId(run.LanguageId) },
                        paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops)
                });
            }
            var saved = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            Assert.False(saved.Report.HasErrors);
            PublisherNativeWritingTests.AssertPreservedStorage(source, saved.Value, "Quill/QuillSub/CONTENTS", "Escher/EscherStm");
            PublisherDocument reopened = PublisherDocument.Load(saved.Value);
            Assert.Equal(masters, reopened.MasterPages.Select(page => page.Id));
            Assert.Equal(sides, reopened.MasterPages.Select(page => page.MasterSide));
            Assert.Equal(assignments, reopened.Pages.Select(page => page.MasterPageId));
            Assert.Equal(document.TextStories.Select(story => story.Text), reopened.TextStories.Select(story => story.Text));
            var frames = reopened.MasterPages.Concat(reopened.Pages).SelectMany(page => page.TextFrames).ToDictionary(frame => frame.Id);
            Assert.Equal(original.Keys.OrderBy(id => id), frames.Keys.OrderBy(id => id));
            foreach (var entry in original) {
                Assert.Equal(entry.Value.X + (iteration + 1) * 6, frames[entry.Key].X, 6);
                Assert.Equal(entry.Value.Y + (iteration + 1) * 12, frames[entry.Key].Y, 6);
                Assert.Equal(entry.Value.Width, frames[entry.Key].Width); Assert.Equal(entry.Value.Height, frames[entry.Key].Height);
            }
            Assert.All(reopened.Pages, page => Assert.Single(page.TextFrames));
            Assert.All(reopened.MasterPages, page => Assert.Single(page.TextFrames));
            Assert.Equal(sides, oldMasters.Select(page => page.MasterSide));
            source = saved.Value;
            document = reopened;
        }
    }
}
