using OfficeIMO;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherDefaultTabTests {
    [Theory]
    [InlineData("Text/native-text.pub", 28.3)]
    [InlineData("TextDefaultTabs/source-sgp72.pub", 72)]
    [InlineData("TextDefaultTabs/source-both36.pub", 36)]
    [InlineData("TextDefaultTabs/source-doc72-empty-sgp.pub", 72)]
    public void Native_default_tab_carriers_survive_unrelated_paragraph_edits(string file, double expected) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + file));
        PublisherDocument document = PublisherDocument.Load(source);
        var story = Assert.Single(document.TextStories);
        var paragraph = Assert.Single(story.Paragraphs);
        Assert.Equal(expected, paragraph.TabStops!.DefaultInterval, 8);
        Assert.Equal(source, document.ToPubBytes());
        var margins = new OfficeTextPadding(paragraph.Margins.Left, paragraph.Margins.Top + 1,
            paragraph.Margins.Right, paragraph.Margins.Bottom);
        var replacement = new OfficeRichTextParagraph(paragraph.Runs, paragraph.Alignment, paragraph.LineHeight,
            margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops);
        document.ReplaceStoryParagraphs(story.Id, new[] { replacement });
        byte[] saved = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        var reloaded = Assert.Single(Assert.Single(PublisherDocument.Load(saved).TextStories).Paragraphs);
        Assert.Equal(expected, reloaded.TabStops!.DefaultInterval, 8);
        Assert.Equal(margins.Top, reloaded.Margins.Top, 8);
        Assert.Equal(paragraph.Runs.Select(run => run.Text), reloaded.Runs.Select(run => run.Text));
    }

    [Fact]
    public void Paragraph_save_refuses_a_changed_document_tab_interval() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TextDefaultTabs/source-sgp72.pub"));
        var story = Assert.Single(document.TextStories);
        var paragraph = Assert.Single(story.Paragraphs);
        Assert.Equal(72, paragraph.TabStops!.DefaultInterval);
        var changed = new OfficeTextTabStops(paragraph.TabStops.Stops, 36, paragraph.TabStops.Origin)
            .WithParagraphAlignment(paragraph.TabStops.AlignWithParagraph);
        document.ReplaceStoryParagraphs(story.Id, new[] {
            new OfficeRichTextParagraph(paragraph.Runs, paragraph.Alignment, paragraph.LineHeight,
                paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(changed)
        });
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(
            new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_EDIT_UNSUPPORTED");
    }
}
