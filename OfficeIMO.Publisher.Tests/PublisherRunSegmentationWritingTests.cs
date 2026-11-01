using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherRunSegmentationWritingTests {
    public static IEnumerable<object[]> SegmentationCases() {
        foreach (bool cells in new[] { false, true })
            foreach (int referencePage in new[] { 1, 2, 3, 4, 5, 7, 8, 9, 10, 13, 15, 16, 17, 18, 19 })
                yield return new object[] { cells, referencePage };
    }

    [Theory]
    [MemberData(nameof(SegmentationCases))]
    public void Retained_paragraph_text_can_split_or_merge_runs_without_moving_native_metadata(bool cells, int referencePage) {
        byte[] source = File.ReadAllBytes(Fixture(cells));
        PublisherDocument original = PublisherDocument.Load(source), document = PublisherDocument.Load(source);
        int page = SourcePage(referencePage);
        PublisherTable? table = cells ? document.Pages[page].Tables.Single() : null;
        PublisherTextStory story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[page];
        var before = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        var expected = cells ? original.Pages[referencePage].Tables.Single().Cells[0].Paragraphs : original.TextStories[referencePage].Paragraphs;
        int[] order = Order(referencePage);
        OfficeRichTextParagraph[] Request(IReadOnlyList<OfficeRichTextParagraph> values) => order.Select((n, i) =>
            Same(values[n], expected[i]) ? values[n] : Compact(expected[i])).ToArray();
        var requested = Request(before);
        Assert.True(order.Where((n, i) => !Same(before[n], expected[i])).Any());
        if (cells) foreach (int index in new[] { 0, 3 }) {
            var cell = table!.Cells[index];
            document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex, Request(cell.Paragraphs));
        } else document.ReplaceStoryParagraphs(story.Id, requested);
        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument saved = PublisherDocument.Load(output);
        var actual = cells ? saved.Pages[page].Tables.Single().Cells[0].Paragraphs : saved.TextStories[page].Paragraphs;
        Assert.Equal(expected.Count, actual.Count);
        byte[] sourceQuill = Quill(source), outputQuill = Quill(output);
        for (int i = 0; i < order.Length; i++) {
            Assert.True(Same(expected[i], actual[i]), $"Independent native final paragraph {i}");
            int length = before[order[i]].Runs.Sum(r => r.Text.Length);
            for (int offset = 0; offset < length; offset++) foreach (byte field in new byte[] { 0x12, 0x3E })
                Assert.Equal(FieldAt(original, sourceQuill, story.Id, order[i], offset, field),
                    FieldAt(saved, outputQuill, story.Id, i, offset, field));
            if (before.Count == requested.Length)
                Assert.Equal(Definition(original, sourceQuill, story.Id, order[i], length), Definition(saved, outputQuill, story.Id, i, length));
        }
        foreach (var other in original.TextStories.Where(s => s.Id != story.Id)) {
            var got = saved.TextStories.Single(s => s.Id == other.Id);
            Assert.Equal(other.Paragraphs.Count, got.Paragraphs.Count);
            Assert.All(other.Paragraphs.Select((p, i) => Same(p, got.Paragraphs[i])), value => Assert.True(value));
        }
        if (cells) foreach (int index in new[] { 1, 2 }) {
            var wanted = original.Pages[page].Tables.Single().Cells[index].Paragraphs;
            var got = saved.Pages[page].Tables.Single().Cells[index].Paragraphs;
            Assert.Equal(wanted.Count, got.Count);
            Assert.All(wanted.Select((p, i) => Same(p, got[i])), value => Assert.True(value));
        }
        PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void Count_changing_run_segmentation_refuses_conflicting_native_language_at_the_same_text_positions(bool cells, bool split, bool finalOnly) {
        PublisherDocument document = PublisherDocument.Load(Fixture(cells));
        int page = split ? 14 : 11;
        PublisherTable? table = cells ? document.Pages[page].Tables.Single() : null;
        PublisherTextStory story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[page];
        var before = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        var reference = cells ? document.Pages[split ? 5 : 1].Tables.Single().Cells[0].Paragraphs[0]
            : document.TextStories[split ? 5 : 1].Paragraphs[0];
        var changed = Compact(reference);
        var requested = finalOnly ? new[] { changed } : new[] { changed, before[2] };
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Run_segmentation_refuses_unsupported_background_changes(bool cells) {
        PublisherDocument document = PublisherDocument.Load(Fixture(cells));
        var table = cells ? document.Pages[0].Tables.Single() : null;
        var story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[0];
        var before = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        var run = new OfficeRichTextRun(Text(before[0]), 12, OfficeColor.Black,
            fontFamily: "Arial", backgroundColor: OfficeColor.Yellow);
        var changed = Copy(before[0], new[] { run });
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, new[] { changed, before[3] });
        else document.ReplaceStoryParagraphs(story.Id, new[] { changed, before[3] });
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    private static int SourcePage(int reference) => reference == 13 ? 12 : reference is >= 7 and <= 10 or 16 or 18 or 19 ? 6 : 0;
    private static int[] Order(int reference) => reference switch {
        1 or 5 or 7 => new[] { 0, 2, 3 }, 2 or 8 => new[] { 3, 1, 0 }, 3 or 9 => new[] { 0, 1, 1, 2, 3 },
        4 or 10 => new[] { 3 }, 13 => new[] { 0, 2 }, 15 or 16 => new[] { 1, 3 }, _ => new[] { 0, 1, 2, 3 }
    };
    private static string Text(OfficeRichTextParagraph p) => string.Concat(p.Runs.Select(r => r.Text));
    private static OfficeRichTextRun At(OfficeRichTextParagraph p, int offset) {
        foreach (var run in p.Runs) { if (offset < run.Text.Length) return run; offset -= run.Text.Length; }
        throw new InvalidOperationException("Position outside paragraph.");
    }
    private static bool Same(OfficeRichTextParagraph a, OfficeRichTextParagraph b) =>
        PublisherTextComparison.ParagraphFormattingEqual(a, b) && Text(a) == Text(b)
        && Enumerable.Range(0, Text(a).Length).All(i => PublisherTextComparison.RunFormattingEqual(At(a, i), At(b, i)));
    private static OfficeRichTextParagraph Compact(OfficeRichTextParagraph p) {
        var runs = new List<OfficeRichTextRun>();
        foreach (var run in p.Runs) {
            if (runs.Count > 0 && PublisherTextComparison.RunFormattingEqual(runs[runs.Count - 1], run)) {
                var previous = runs[runs.Count - 1]; runs[runs.Count - 1] = WithText(previous, previous.Text + run.Text);
            } else runs.Add(run);
        }
        return Copy(p, runs);
    }
    private static OfficeRichTextRun WithText(OfficeRichTextRun r, string text) => new OfficeRichTextRun(text, r.FontSize, r.Color, r.Bold, r.Italic,
        r.Underline, r.FontFamily, r.Strikethrough, r.BackgroundColor, r.UnderlineStyle, r.StrikethroughStyle, r.Baseline, r.ParagraphIndent)
        { LinkUri = r.LinkUri, LinkDestinationName = r.LinkDestinationName };
    private static OfficeRichTextParagraph Copy(OfficeRichTextParagraph p, IReadOnlyList<OfficeRichTextRun> runs) =>
        new OfficeRichTextParagraph(runs, p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
    private static byte[] Quill(byte[] source) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        return compound!.Streams["Quill/QuillSub/CONTENTS"];
    }
    private static byte[] Definition(PublisherDocument d, byte[] quill, uint id, int paragraph, int offset) {
        var binding = d.SourceStyleCatalog.Stories.Single(s => s.StoryId == id);
        int position = binding.Uses.Where(u => u.Tag == "FDPP").ElementAt(paragraph).NativeStart + offset;
        var span = binding.Uses.Single(u => u.Tag == "FDPC" && u.NativeStart <= position && u.NativeStart + u.NativeLength > position).DirectDefinition;
        return quill.Skip(span.Offset).Take(span.Length).ToArray();
    }
    private static byte[]? FieldAt(PublisherDocument d, byte[] quill, uint id, int paragraph, int offset, byte field) {
        byte[] raw = Definition(d, quill, id, paragraph, offset);
        var data = new PublisherBinaryData(raw, "test/native-character-definition");
        var fields = new PublisherBlockReader(data, new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None))
            .Chunk(0, raw.Length).Where(f => f.Id == field).ToArray();
        if (fields.Length == 0) return null;
        var match = Assert.Single(fields); return raw.Skip(match.Offset - 2).Take(match.Length + 2).ToArray();
    }
    private static string Fixture(bool cells) => PublisherNativeTests.Fixture("NativeComparison/"
        + (cells ? "TableRunSegmentation/native-table" : "TextRunSegmentation/native-text") + "runsegmentation.pub");
}
