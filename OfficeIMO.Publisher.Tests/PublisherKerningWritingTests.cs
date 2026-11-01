using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherKerningWritingTests {
    [Fact]
    public void Threshold_metadata_survives_authored_linked_projection_and_run_copy_methods() {
        PublisherDocument document = PublisherDocument.Create(300, 400);
        var run = new OfficeRichTextRun(string.Concat(Enumerable.Repeat("Complete threshold metadata. ", 40)), 12,
            OfficeColor.Black, fontFamily: "Arial").WithAutomaticPairKerningThreshold(8).WithLanguageId(1033).WithParagraphIndent(OfficeTextParagraphIndent.Empty);
        PublisherTextFrame first = document.AddTextFrame(document.Pages[0].Id, 20, 20, 120, 60,
            new[] { new OfficeRichTextParagraph(new[] { run }, lineHeight: 18) });
        PublisherPage page = document.AddPage();
        document.AddLinkedTextFrame(first.Id, page.Id, 20, 20, 120, 60);
        Assert.True(document.Pages[1].TextFrames.Single().HasOverflow);
        Assert.Equal(run.Text, Assert.Single(document.TextStories).Text);
        Assert.All(Assert.Single(document.TextStories).Paragraphs.SelectMany(paragraph => paragraph.Runs),
            value => Assert.Equal(8D, value.AutomaticPairKerningThreshold));
        OfficeDrawingRichText[] painted = document.Pages.SelectMany(item => PublisherNativeTests.Elements(item.CreateDrawing())).OfType<OfficeDrawingRichText>().ToArray();
        Assert.Equal(2, painted.Length);
        Assert.All(painted.SelectMany(item => item.Paragraphs).SelectMany(paragraph => paragraph.Runs),
            value => Assert.Equal(8D, value.AutomaticPairKerningThreshold));
        Assert.Equal(8D, run.AutomaticPairKerningThreshold);
        Assert.Equal(0D, run.WithAutomaticPairKerningThreshold(0).AutomaticPairKerningThreshold);
        Assert.Equal(8D, run.AutomaticPairKerningThreshold);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(8)]
    [InlineData(12)]
    [InlineData(14)]
    public void Native_resaved_threshold_controls_recover_stored_metadata_and_copy_without_changes(int threshold) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(
            "NativeComparison/TextPairKerning/native-threshold-" + threshold + ".pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(15, document.Pages.Count);
        // Publisher removes the observed 14-point default from these direct definitions.
        // Keep omitted/default metadata distinct from an explicit stored value.
        Assert.All(document.TextStories.SelectMany(story => story.Paragraphs).SelectMany(paragraph => paragraph.Runs),
            run => Assert.Equal(threshold == 14 ? (double?)null : threshold, run.AutomaticPairKerningThreshold));
        Assert.Equal(source, document.ToPubBytes());
    }

    [Fact]
    public void Explicit_threshold_set_reset_and_unspecified_saves_keep_text_opaque_fields_and_neighboring_stories() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TextPairKerning/native-threshold-8.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        uint id = document.TextStories.First().Id;
        var stories = document.TextStories.Select(story => (story.Id, story.Text)).ToArray();
        var frames = document.Pages.SelectMany(page => page.TextFrames).Select(frame => (frame.Id, frame.StoryId, frame.PageId)).ToArray();
        foreach (double? threshold in new double?[] { 0, null, 8.2, 12.5, 999.5, 14 }) {
            document.ReplaceStoryParagraphs(id, document.TextStories.Single(story => story.Id == id).Paragraphs
                .Select(paragraph => WithThreshold(paragraph, threshold)).ToArray());
            byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");
            Assert.Equal(OpaqueCharacterFields(source), OpaqueCharacterFields(output));
            document = PublisherDocument.Load(output);
            Assert.Equal(stories, document.TextStories.Select(story => (story.Id, story.Text)));
            Assert.Equal(frames, document.Pages.SelectMany(page => page.TextFrames).Select(frame => (frame.Id, frame.StoryId, frame.PageId)));
            Assert.All(document.TextStories.Single(story => story.Id == id).Paragraphs.SelectMany(paragraph => paragraph.Runs),
                run => Assert.Equal(threshold ?? 0, run.AutomaticPairKerningThreshold));
            Assert.All(document.TextStories.Where(story => story.Id != id).SelectMany(story => story.Paragraphs).SelectMany(paragraph => paragraph.Runs),
                run => Assert.Equal(8D, run.AutomaticPairKerningThreshold));
            source = output;
        }
    }

    [Fact]
    public void Fractional_native_point_values_transfer_between_runs_without_binary_rounding_rejection() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"));
        PublisherTextStory story = document.TextStories[0];
        var template = story.Paragraphs[0];
        var original = template.Runs[0];
        var run = new OfficeRichTextRun(original.Text, 8.2, original.Color, fontFamily: original.FontFamily)
            .WithAutomaticPairKerningThreshold(8.2);
        document.ReplaceStoryParagraphs(story.Id, new[] { new OfficeRichTextParagraph(new[] { run }, template.Alignment,
            template.LineHeight, template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops) });
        document = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        var recovered = Assert.Single(Assert.Single(document.TextStories).Paragraphs).Runs[0];
        Assert.Equal(8.2, recovered.FontSize);
        Assert.Equal(8.2, recovered.AutomaticPairKerningThreshold);
        // Transfer a decoded value to a fresh source definition, requiring encoding again.
        document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"));
        document.ReplaceStoryParagraphs(story.Id, new[] { new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("Transferred native points", recovered.FontSize, recovered.Color, fontFamily: recovered.FontFamily)
                .WithAutomaticPairKerningThreshold(recovered.AutomaticPairKerningThreshold)
        }, template.Alignment, template.LineHeight, template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops) });
        var saved = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains("Transferred native points", Assert.Single(saved.TextStories).Text);
        var transferred = Assert.Single(Assert.Single(saved.TextStories).Paragraphs).Runs[0];
        Assert.Equal(8.2, transferred.FontSize);
        Assert.Equal(8.2, transferred.AutomaticPairKerningThreshold);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Retained_paragraph_count_changes_apply_threshold_with_or_without_text_replacement(bool cells, bool replaceText) {
        string leaf = cells ? "TableNonuniformParagraphs/native-tablenonuniformparagraphs.pub" : "TextNonuniformParagraphs/native-textnonuniformparagraphs.pub";
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/" + leaf));
        PublisherTable? table = cells ? document.Pages[0].Tables.Single() : null;
        PublisherTextStory story = cells ? document.TextStories.Single(value => value.Id == table!.StoryId) : document.TextStories[0];
        var before = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        var requested = before.Skip(1).Select(paragraph => WithThreshold(paragraph, 8, replaceText)).ToArray();
        Assert.NotEmpty(requested);
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        PublisherDocument saved = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        var actual = cells ? saved.Pages[0].Tables.Single().Cells[0].Paragraphs : saved.TextStories[0].Paragraphs;
        Assert.Equal(requested.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))),
            actual.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));
        Assert.All(actual.SelectMany(paragraph => paragraph.Runs), run => Assert.Equal(8D, run.AutomaticPairKerningThreshold));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Retained_empty_typing_CR_receives_explicit_threshold_when_topology_grows(bool cells) {
        string leaf = cells ? "TableCellTyping/native-tablecelltyping.pub" : "TextUniformEmptyParagraphs/native-textuniformemptyparagraphs.pub";
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/" + leaf));
        if (!cells) {
            PublisherTextStory source = document.TextStories[0];
            document.ReplaceStoryParagraphs(source.Id, new[] { source.Paragraphs[0] });
            document = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        }
        PublisherTable? table = cells ? document.Pages.SelectMany(page => page.Tables).First(item =>
            item.Cells.Any(cell => cell.Paragraphs.Count == 1 && cell.Paragraphs[0].Runs.All(run => run.Text.Length == 0))) : null;
        PublisherTableCell? cell = table?.Cells.First(value => value.Paragraphs.Count == 1 && value.Paragraphs[0].Runs.All(run => run.Text.Length == 0));
        PublisherTextStory story = cells ? document.TextStories.Single(value => value.Id == table!.StoryId)
            : document.TextStories.First(value => value.Paragraphs.Count == 1 && value.Paragraphs[0].Runs.All(run => run.Text.Length == 0));
        OfficeRichTextParagraph original = cells ? cell!.Paragraphs[0] : story.Paragraphs.Single();
        var typed = WithThreshold(original, 8);
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, cell!.RowIndex, cell.ColumnIndex, new[] { typed, typed });
        else document.ReplaceStoryParagraphs(story.Id, new[] { typed, typed });
        PublisherDocument saved = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        var actual = cells ? saved.Pages.SelectMany(page => page.Tables).Single(value => value.Id == table!.Id)
            .Cells.Single(value => value.RowIndex == cell!.RowIndex && value.ColumnIndex == cell.ColumnIndex).Paragraphs
            : saved.TextStories.Single(value => value.Id == story.Id).Paragraphs;
        Assert.Equal(2, actual.Count);
        Assert.All(actual.SelectMany(paragraph => paragraph.Runs), run => { Assert.Equal(string.Empty, run.Text); Assert.Equal(8D, run.AutomaticPairKerningThreshold); });
    }

    [Theory]
    [InlineData(999.6)]
    [InlineData(0.00001)]
    public void Unrepresentable_threshold_requests_remain_refused_when_loss_is_allowed(double threshold) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"));
        PublisherTextStory story = document.TextStories[0];
        document.ReplaceStoryParagraphs(story.Id, story.Paragraphs.Select(paragraph => WithThreshold(paragraph, threshold)).ToArray());
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unqualified_threshold_source_stays_opaque_and_refuses_explicit_rewriting(bool outOfRange) {
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            var data = new PublisherBinaryData(bytes, "Quill");
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            PublisherQuillChunk page = PublisherQuillDirectory.Read(data, context).First(chunk => chunk.Tag == "FDPC");
            int count = data.U16(page.Offset);
            int position = page.Offset + data.U16(page.Offset + 8 + count * 4);
            PublisherBlock field = new PublisherBlockReader(data, context).Chunk(position, page.End).Single(value => value.Id == 0x18);
            if (outOfRange) PublisherInputContractTests.WriteUInt32(bytes, field.Offset, uint.MaxValue);
            else bytes[field.Offset - 1] = 0x20;
        }, "NativeComparison/TextPairKerning/native-threshold-8.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_TEXT_KERNING_UNASSESSED");
        Assert.Equal(source, document.ToPubBytes());
        PublisherTextStory story = document.TextStories[0];
        document.ReplaceStoryParagraphs(story.Id, story.Paragraphs.Select(paragraph => WithThreshold(paragraph, 0)).ToArray());
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    private static OfficeRichTextParagraph WithThreshold(OfficeRichTextParagraph paragraph, double? threshold, bool replaceText = false) =>
        new OfficeRichTextParagraph(paragraph.Runs.Select(run => {
            var copy = new OfficeRichTextRun(replaceText ? run.Text + " extra" : run.Text, run.FontSize, run.Color, run.Bold,
                run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle,
                run.StrikethroughStyle, run.Baseline, run.ParagraphIndent).WithLanguageId(run.LanguageId);
            return copy.WithAutomaticPairKerningThreshold(threshold);
        }).ToArray(), paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor)
            .WithTabStops(paragraph.TabStops);

    private static string[] OpaqueCharacterFields(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        var data = new PublisherBinaryData(compound!.Streams["Quill/QuillSub/CONTENTS"], "Quill");
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        var blocks = new PublisherBlockReader(data, context);
        return PublisherQuillDirectory.Read(data, context).Where(chunk => chunk.Tag == "FDPC").SelectMany(chunk => {
            int count = data.U16(chunk.Offset);
            return Enumerable.Range(0, count).Select(index => {
                int position = chunk.Offset + data.U16(chunk.Offset + 8 + count * 4 + index * 2);
                return string.Join("|", blocks.Chunk(position, chunk.End).Where(field => field.Id != 0x18)
                    .Select(field => Convert.ToBase64String(data.Bytes, field.Offset - 2, field.Length + 2)));
            });
        }).ToArray();
    }
}
