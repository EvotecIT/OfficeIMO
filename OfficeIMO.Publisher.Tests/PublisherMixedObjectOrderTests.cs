using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherMixedObjectOrderTests {
    [Fact]
    public void Mixed_additions_preserve_append_order_and_old_page_scenes() {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        uint pageId = document.Pages[0].Id;
        var paint = OfficeShape.Rectangle(240, 120);
        paint.FillColor = OfficeColor.Red; paint.StrokeColor = null;
        PublisherShape background = document.AddShape(pageId, paint, 72, 144);
        PublisherPicture picture = document.AddPicture(pageId,
            PublisherPictureEditingFixture.Image(OfficeColor.Blue), 108, 144, 96, 96);
        var oldPage = document.Pages[0];
        PublisherTextFrame frame = document.AddTextFrame(pageId, 72, 144, 288, 108,
            new[] { new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Overlapping text", 18, OfficeColor.Black) }) });
        paint.FillColor = OfficeColor.Green;
        PublisherShape foreground = document.AddShape(pageId, paint, 72, 168);
        Assert.Equal(new[] { background.Id, picture.Id, frame.Id, foreground.Id }, Order(document.Pages[0]));
        Assert.Equal(new[] { background.Id, picture.Id }, Order(oldPage));
        document.SetPictureBounds(picture.Id, 120, 144, 96, 96);
        document.ReplaceStoryParagraphs(frame.StoryId, new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Changed text", 18, OfficeColor.Black) })
        });
        Assert.Equal(new[] { background.Id, picture.Id, frame.Id, foreground.Id }, Order(document.Pages[0]));
        Assert.Equal(new[] { background.Id, picture.Id }, Order(oldPage));
    }

    [Theory]
    [InlineData(0, new uint[] { 299, 300, 302, 303 })]
    [InlineData(2, new uint[] { 304, 305, 307, 308 })]
    public void Native_overlap_control_projects_saved_cross_kind_order(int pageIndex, uint[] expected) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/MixedObjectPages/native-mixedobjectpages.pub"));
        Assert.Empty(Order(document.Pages[1]));
        Assert.Equal(expected, Order(document.Pages[pageIndex]));
        Assert.Equal(3, document.Pages.SelectMany(page => page.Pictures).Count());
        Assert.Single(document.Images);
    }

    private static uint[] Order(PublisherPage page) => PublisherNativeTests.Elements(page.CreateDrawing())
        .SelectMany(element => element.SourceElementIds ?? Array.Empty<string>())
        .Where(id => id.StartsWith("publisher-object-", StringComparison.Ordinal))
        .Select(id => uint.Parse(id.Substring("publisher-object-".Length), System.Globalization.CultureInfo.InvariantCulture))
        .Distinct().ToArray();

    [Fact]
    public void Generated_compact_stories_can_be_edited_before_native_resave() {
        PublisherDocument document = Generated();
        PublisherTextStory first = document.TextStories.Single(story => story.Id == 1);
        string other = document.TextStories.Single(story => story.Id == 2).Text;
        uint[][] order = document.Pages.Select(Order).ToArray();
        byte[] image = Assert.Single(document.Images).GetBytes();
        document.ReplaceStoryParagraphs(first.Id, new[] { Replacement(first.Paragraphs.Single(), "Edited first: overlap order") });
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        PublisherDocument reopened = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Equal("Edited first: overlap order\n", reopened.TextStories.Single(story => story.Id == first.Id).Text);
        Assert.Equal(other, reopened.TextStories.Single(story => story.Id == 2).Text);
        Assert.Equal(image, Assert.Single(reopened.Images).GetBytes());
        for (int page = 0; page < order.Length; page++) Assert.Equal(order[page], Order(reopened.Pages[page]));
    }

    [Theory]
    [InlineData("color")]
    [InlineData("font")]
    public void Compact_source_refuses_unqualified_resource_growth_without_publishing_output(string resource) {
        PublisherDocument document = Generated();
        PublisherTextStory story = document.TextStories.Single(item => item.Id == 1);
        OfficeRichTextParagraph before = story.Paragraphs.Single();
        OfficeRichTextParagraph after = Replacement(before, "Edited first: overlap order",
            resource == "color" ? OfficeColor.Blue : null, resource == "font" ? "Calibri" : null);
        document.ReplaceStoryParagraphs(story.Id, new[] { after });
        using var output = new MemoryStream();
        OfficeConversionException failure = Assert.Throws<OfficeConversionException>(() => document.Save(output,
            new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(failure.Report.FidelityDiagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Failure);
        Assert.Equal(0, output.Length);
        OfficeRichTextRun retained = document.TextStories.Single(item => item.Id == 1).Paragraphs.Single().Runs.Single();
        Assert.Equal("Edited first: overlap order", retained.Text);
        if (resource == "color") Assert.Equal(OfficeColor.Blue, retained.Color);
        else Assert.Equal("Calibri", retained.FontFamily);
    }

    private static PublisherDocument Generated() => PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/GeneratedMixedObjectPages/generated-mixedobjectpages.pub"));

    [Theory]
    [InlineData(false, 100, 1)]
    [InlineData(false, 1800, 2)]
    [InlineData(true, 900, 2)]
    public void Compact_formatting_and_directory_growth_remains_editable_on_the_next_save(bool paragraphs, int count, int minimumDirectories) {
        PublisherDocument document = Generated();
        OfficeRichTextParagraph original = document.TextStories.Single(story => story.Id == 1).Paragraphs.Single();
        string other = document.TextStories.Single(story => story.Id == 2).Text;
        byte[] image = Assert.Single(document.Images).GetBytes();
        uint[][] order = document.Pages.Select(Order).ToArray();
        document.ReplaceStoryParagraphs(1, Content(false));
        byte[] first = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.True(OfficeIMO.Core.Internal.OfficeCompoundFileReader.TryRead(first,
            out OfficeIMO.Core.Internal.OfficeCompoundFile? compound, out string? error), error);
        byte[] quill = compound!.Streams["Quill/QuillSub/CONTENTS"];
        Assert.True(BitConverter.ToUInt32(quill, 16) >= minimumDirectories * 512U);
        PublisherDocument reopened = PublisherDocument.Load(first);
        reopened.ReplaceStoryParagraphs(1, Content(true));
        PublisherDocument final = PublisherDocument.Load(reopened.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        PublisherTextStory current = final.TextStories.Single(story => story.Id == 1);
        Assert.Equal(paragraphs ? count : 1, current.Paragraphs.Count);
        Assert.Equal(paragraphs ? "B\n" + string.Concat(Enumerable.Repeat("A\n", count - 1)) : "B" + new string('A', count - 1) + "\n", current.Text);
        var runs = current.Paragraphs.SelectMany(paragraph => paragraph.Runs).ToArray();
        Assert.Equal(count, runs.Length);
        for (int index = 0; index < count; index++) {
            Assert.Equal(!paragraphs && index % 2 == 1, runs[index].Bold);
            Assert.Equal(18, runs[index].FontSize);
        }
        Assert.Equal(other, final.TextStories.Single(story => story.Id == 2).Text);
        Assert.Equal(image, Assert.Single(final.Images).GetBytes());
        for (int page = 0; page < order.Length; page++) Assert.Equal(order[page], Order(final.Pages[page]));

        OfficeRichTextParagraph[] Content(bool again) => paragraphs
            ? Enumerable.Range(0, count).Select(index => Replacement(original, index == 0 && again ? "B" : "A")).ToArray()
            : new[] { new OfficeRichTextParagraph(Enumerable.Range(0, count).Select(index => new OfficeRichTextRun(
                index == 0 && again ? "B" : "A", 18, OfficeColor.Black, bold: index % 2 == 1, fontFamily: "Arial")).ToArray(),
                original.Alignment, original.LineHeight, original.Margins, original.Indent, original.LineHeightFactor).WithTabStops(original.TabStops) };
    }

    private static OfficeRichTextParagraph Replacement(OfficeRichTextParagraph paragraph, string text,
        OfficeColor? color = null, string? font = null) {
        OfficeRichTextRun run = paragraph.Runs.Single();
        return new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun(text, run.FontSize, color ?? run.Color, run.Bold, run.Italic, run.Underline,
                font ?? run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle,
                run.StrikethroughStyle, run.Baseline, run.ParagraphIndent)
        }, paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops);
    }
}
