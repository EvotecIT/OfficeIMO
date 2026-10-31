using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherEmptyRefillWritingTests {
    public static IEnumerable<object[]> RefillCases() {
        foreach (bool cells in new[] { false, true })
            foreach (int page in new[] { 1, 2, 3, 4, 5, 6, 9 }) yield return new object[] { cells, page };
    }

    [Theory]
    [MemberData(nameof(RefillCases))]
    public void Empty_native_typing_can_fill_paragraphs_during_nonuniform_count_changes(bool cells, int referencePage) {
        byte[] source = File.ReadAllBytes(Fixture(cells));
        PublisherDocument original = PublisherDocument.Load(source), document = PublisherDocument.Load(source);
        int sourcePage = referencePage == 9 ? 8 : 0;
        PublisherTable? table = cells ? document.Pages[sourcePage].Tables.Single() : null;
        PublisherTextStory story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[sourcePage];
        int[] order = referencePage switch {
            1 => new[] { 1, 2, 3 }, 2 => new[] { 0, 1, 3 }, 3 => new[] { 3 }, 4 => new[] { 1 },
            5 => new[] { 3, 0, 1 }, 6 => new[] { 0, 1, 1, 2, 3 }, _ => new[] { 0, 2 }
        };
        var expected = cells ? original.Pages[referencePage].Tables.Single().Cells[0].Paragraphs : original.TextStories[referencePage].Paragraphs;
        var selected = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        OfficeRichTextParagraph[] Fill(IReadOnlyList<OfficeRichTextParagraph> paragraphs) => order.Select((index, i) =>
            string.Concat(paragraphs[index].Runs.Select(r => r.Text)) == string.Concat(expected[i].Runs.Select(r => r.Text))
                ? paragraphs[index] : WithText(paragraphs[index], string.Concat(expected[i].Runs.Select(r => r.Text)))).ToArray();
        if (cells) foreach (int index in new[] { 0, 3 }) {
            PublisherTableCell cell = table!.Cells[index];
            document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex, Fill(cell.Paragraphs));
        } else document.ReplaceStoryParagraphs(story.Id, Fill(selected));
        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument saved = PublisherDocument.Load(output);
        var actual = cells ? saved.Pages[sourcePage].Tables.Single().Cells[0].Paragraphs : saved.TextStories[sourcePage].Paragraphs;
        Assert.Equal(expected.Count, actual.Count);
        for (int i = 0; i < order.Length; i++) {
            Assert.True(PublisherTextComparison.Paragraphs.Equals(expected[i], actual[i]), $"Native final paragraph {i}");
            if (selected[order[i]].Runs.Single().Text.Length == 0 && actual[i].Runs.Single().Text.Length != 0) {
                Assert.Equal(Definition(original, source, story.Id, order[i], true), Definition(saved, output, story.Id, i, false));
                Assert.Equal(Definition(original, source, story.Id, order[i], true), Definition(saved, output, story.Id, i, true));
            }
        }
        foreach (PublisherTextStory other in original.TextStories.Where(s => s.Id != story.Id))
            Assert.True(other.Paragraphs.SequenceEqual(saved.TextStories.Single(s => s.Id == other.Id).Paragraphs, PublisherTextComparison.Paragraphs));
        if (cells) foreach (int index in new[] { 1, 2 })
            Assert.True(original.Pages[sourcePage].Tables.Single().Cells[index].Paragraphs.SequenceEqual(
                saved.Pages[sourcePage].Tables.Single().Cells[index].Paragraphs, PublisherTextComparison.Paragraphs));
        PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Empty_refill_refuses_same_visible_typing_with_conflicting_native_language(bool cells, bool finalOnly) {
        PublisherDocument document = PublisherDocument.Load(Fixture(cells));
        PublisherTable? table = cells ? document.Pages[7].Tables.Single() : null;
        PublisherTextStory story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[7];
        var paragraphs = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        var requested = finalOnly ? new[] { WithText(paragraphs[0], "Filled conflicting template") }
            : new[] { WithText(paragraphs[0], "Filled conflicting template"), paragraphs[2] };
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    private static OfficeRichTextParagraph WithText(OfficeRichTextParagraph paragraph, string text) {
        OfficeRichTextRun run = Assert.Single(paragraph.Runs);
        Assert.Empty(run.Text);
        return new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, run.FontSize, run.Color, run.Bold, run.Italic,
            run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle,
            run.StrikethroughStyle, run.Baseline, run.ParagraphIndent) }, paragraph.Alignment, paragraph.LineHeight,
            paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops);
    }

    private static string Fixture(bool cells) => PublisherNativeTests.Fixture("NativeComparison/"
        + (cells ? "TableEmptyRefill/native-table" : "TextEmptyRefill/native-text") + "emptyrefill.pub");

    private static byte[] Definition(PublisherDocument document, byte[] source, uint id, int paragraph, bool cr) {
        var binding = document.SourceStyleCatalog.Stories.Single(s => s.StoryId == id);
        var pp = binding.Uses.Where(u => u.Tag == "FDPP").ElementAt(paragraph);
        int position = pp.NativeStart + (cr ? pp.NativeLength - 1 : 0);
        var span = binding.Uses.Single(u => u.Tag == "FDPC" && u.NativeStart <= position && u.NativeStart + u.NativeLength > position).DirectDefinition;
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        return compound!.Streams["Quill/QuillSub/CONTENTS"].Skip(span.Offset).Take(span.Length).ToArray();
    }
}
