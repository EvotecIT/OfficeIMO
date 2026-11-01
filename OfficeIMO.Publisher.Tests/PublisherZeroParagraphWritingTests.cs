using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherZeroParagraphWritingTests {
    [Theory]
    [InlineData("NativeComparison/Text/native-text.pub")]
    [InlineData("NativeComparison/TextRunSegmentation/native-textrunsegmentation.pub")]
    [InlineData("NativeComparison/LinkedStory/native-linkedstory.pub")]
    [InlineData("NativeComparison/TextStories/native-textstories.pub")]
    [InlineData("native-cleared-text.pub")]
    public void Empty_public_stories_save_reported_native_placeholders_without_mutating_the_model(string fixture) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(fixture));
        var original = PublisherDocument.Load(source); var d = PublisherDocument.Load(source); uint id = d.TextStories[0].Id;
        d.ReplaceStoryParagraphs(id, Array.Empty<OfficeRichTextParagraph>()); var snapshot = d.Pages[0].TextFrames[0];
        var result = d.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        var diagnostic = Assert.Single(result.Report.FidelityDiagnostics, x => x.Code == "PUB_NATIVE_EMPTY_STORY_PARAGRAPH");
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.False(result.Report.HasErrors); Assert.True(result.Report.HasLoss);
        Assert.Empty(d.TextStories.Single(s => s.Id == id).Paragraphs); Assert.Empty(d.TextStories.Single(s => s.Id == id).Text);
        Assert.Same(snapshot, d.Pages[0].TextFrames[0]);
        var saved = PublisherDocument.Load(result.Value); var story = saved.TextStories.Single(s => s.Id == id);
        Assert.Equal("\n", story.Text); var paragraph = Assert.Single(story.Paragraphs); Assert.Empty(Assert.Single(paragraph.Runs).Text);
        Assert.Equal(Definition(original, source, id, "FDPP"), Definition(saved, result.Value, id, "FDPP"));
        Assert.Equal(Definition(original, source, id, "FDPC"), Definition(saved, result.Value, id, "FDPC"));
        foreach (var neighbor in original.TextStories.Where(s => s.Id != id)) Assert.Equal(neighbor.Text, saved.TextStories.Single(s => s.Id == neighbor.Id).Text);
        if (fixture != "native-cleared-text.pub") PublisherNativeWritingTests.AssertPreservedStorage(source, result.Value, "Quill/QuillSub/CONTENTS");
        Assert.Equal(result.Value, saved.ToPubBytes()); Assert.Equal(result.Value, d.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        if (fixture == "native-cleared-text.pub") { Assert.True(result.Report.IsExactSourceCopy); Assert.Equal(source, result.Value); }
        var refill = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Refilled after native clear", 12, OfficeColor.Blue, fontFamily: "Arial") },
            paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops);
        saved.ReplaceStoryParagraphs(id, new[] { refill });
        var reloaded = PublisherDocument.Load(saved.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(PublisherTextComparison.ParagraphEqual(refill, Assert.Single(reloaded.TextStories.Single(s => s.Id == id).Paragraphs)));
    }

    [Fact]
    public void Native_already_empty_source_still_requires_allow_for_the_reported_paragraph_mapping() {
        var d = PublisherDocument.Load(PublisherNativeTests.Fixture("native-cleared-text.pub"));
        d.ReplaceStoryParagraphs(d.TextStories[0].Id, Array.Empty<OfficeRichTextParagraph>());
        using var destination = new MemoryStream(); destination.Write(new byte[] { 9, 8, 7, 6 }, 0, 4); destination.Position = 2;
        var error = Assert.Throws<OfficeConversionException>(() => d.Save(destination));
        var report = Assert.IsType<PublisherSaveReport>(error.Report); Assert.False(report.HasErrors);
        Assert.Contains(report.FidelityDiagnostics, x => x.Code == "PUB_NATIVE_EMPTY_STORY_PARAGRAPH" && x.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Equal(new byte[] { 9, 8, 7, 6 }, destination.ToArray()); Assert.Equal(2, destination.Position); Assert.True(destination.CanWrite);
        Assert.Empty(d.TextStories[0].Paragraphs);
    }

    [Fact]
    public void Clearing_does_not_erase_unsupported_native_label_definitions() {
        var d = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TextLabelFitting/native-textlabelfitting.pub"));
        var story = d.TextStories[0]; Assert.Contains(story.Paragraphs, p => p.Label != null);
        d.ReplaceStoryParagraphs(story.Id, Array.Empty<OfficeRichTextParagraph>());
        var error = Assert.Throws<OfficeConversionException>(() => d.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
        Assert.DoesNotContain(error.Report.FidelityDiagnostics, x => x.Code == "PUB_NATIVE_EMPTY_STORY_PARAGRAPH");
    }

    private static byte[] Definition(PublisherDocument d, byte[] source, uint id, string tag) {
        var b = d.SourceStyleCatalog.Stories.Single(s => s.StoryId == id); var pp = b.Uses.Last(u => u.Tag == "FDPP");
        int end = pp.NativeStart + pp.NativeLength - 1;
        var span = tag == "FDPP" ? pp.DirectDefinition : b.Uses.Single(u => u.Tag == "FDPC" && u.NativeStart <= end && u.NativeStart + u.NativeLength > end).DirectDefinition;
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        return compound!.Streams["Quill/QuillSub/CONTENTS"].Skip(span.Offset).Take(span.Length).ToArray();
    }
}
