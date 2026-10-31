using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherEmptyParagraphRetentionWritingTests {
    [Theory]
    [InlineData(false, 0, false, 0)]
    [InlineData(false, 1, false, 0)]
    [InlineData(false, 0, true, 0)]
    [InlineData(true, 0, false, 0)]
    [InlineData(true, 1, false, 0)]
    [InlineData(true, 0, true, 0)]
    [InlineData(false, 0, false, 1)]
    [InlineData(false, 1, false, 2)]
    [InlineData(false, 0, true, 1)]
    [InlineData(false, 0, true, 2)]
    [InlineData(true, 0, false, 1)]
    [InlineData(true, 1, false, 2)]
    [InlineData(true, 0, true, 1)]
    [InlineData(true, 0, true, 2)]
    public void Uniform_empty_paragraphs_resolve_native_identity_before_the_topology_shortcut(bool cells, int selected, bool copied, int authored) {
        string leaf = cells ? "TableUniformEmptyParagraphs/native-tableuniformemptyparagraphs.pub" : "TextUniformEmptyParagraphs/native-textuniformemptyparagraphs.pub";
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + leaf));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable? table = cells ? document.Pages[0].Tables.Single() : null;
        PublisherTextStory story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[0];
        var paragraphs = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        Assert.Equal(2, paragraphs.Count);
        Assert.NotEqual(Cr(document, source, story.Id, 0), Cr(document, source, story.Id, 1));
        OfficeRichTextParagraph p = paragraphs[selected];
        OfficeRichTextParagraph retained = copied ? p.WithTabStops(p.TabStops) : p;
        OfficeRichTextParagraph Text(string text) => new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun(text, p.Runs[0].FontSize, p.Runs[0].Color, p.Runs[0].Bold, p.Runs[0].Italic, fontFamily: p.Runs[0].FontFamily)
        }, p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
        var requested = authored == 0 ? new[] { retained } : authored == 1
            ? new[] { Text("First authored paragraph"), Text("Second authored paragraph"), retained }
            : new[] { retained, Text("First authored paragraph"), Text("Second authored paragraph") };
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        if (copied) {
            Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        } else {
            byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            PublisherDocument saved = PublisherDocument.Load(output);
            Assert.Equal(Cr(document, source, story.Id, selected), Cr(saved, output, story.Id, authored == 1 ? 2 : 0));
            PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");
        }
    }

    [Theory]
    [InlineData(11, 12)]
    [InlineData(12, 10)]
    [InlineData(13, 14)]
    [InlineData(14, 12)]
    [InlineData(15, 14)]
    public void Native_western_font_attributes_use_snapshot_defaults_and_remain_distinct_from_RTL_attributes(int page, double points) {
        PublisherDocument document = PublisherDocument.Load(Fixture(false));
        OfficeRichTextRun run = Assert.Single(Assert.Single(document.TextStories[page].Paragraphs).Runs);
        Assert.Equal("Native style size witness", run.Text); Assert.Equal("Arial", run.FontFamily); Assert.Equal(points, run.FontSize);
        Assert.Equal(page == 14, run.Bold); Assert.Equal(page == 14, run.Italic);
    }

    public static IEnumerable<object[]> RetentionCases() {
        foreach (bool cells in new[] { false, true })
            foreach (string operation in new[] { "drop-first", "drop-middle", "drop-final", "reorder", "duplicate", "detached",
                "language-0", "language-1", "interior-last", "only-final", "two-empty", "empty-language-0", "empty-language-1" })
                yield return new object[] { cells, operation };
    }

    [Theory]
    [MemberData(nameof(RetentionCases))]
    public void Retained_empty_paragraphs_keep_their_own_native_typing_and_layout_across_count_changes(bool cells, string operation) {
        byte[] source = File.ReadAllBytes(Fixture(cells)); PublisherDocument document = PublisherDocument.Load(source);
        int page = operation.StartsWith("empty-language", StringComparison.Ordinal) ? 10
            : operation.StartsWith("language", StringComparison.Ordinal) ? 6 : 0;
        var originals = document.TextStories.ToArray();
        PublisherTable? table = cells ? document.Pages[page].Tables.Single() : null;
        PublisherTextStory story = cells ? originals.Single(s => s.Id == table!.StoryId) : originals[page];
        var paragraphs = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        if (page == 0) {
            Assert.Equal("Arial", paragraphs[1].Runs[0].FontFamily); Assert.Equal(12, paragraphs[1].Runs[0].FontSize);
            Assert.Equal(cells ? "Arial" : "Calibri", paragraphs[3].Runs[0].FontFamily);
            Assert.Equal(cells ? 12 : 10, paragraphs[3].Runs[0].FontSize);
        }
        int[] order = operation switch {
            "drop-first" => new[] { 1, 2, 3 }, "drop-middle" => new[] { 0, 2, 3 }, "drop-final" => new[] { 0, 1, 2 },
            "reorder" or "detached" => new[] { 3, 0, 2 }, "duplicate" => new[] { 0, 1, 1, 2, 3 },
            "language-0" or "empty-language-0" => new[] { 0, 2 }, "language-1" or "empty-language-1" => new[] { 1, 2 },
            "interior-last" => new[] { 0, 1 }, "only-final" => new[] { 3 }, _ => new[] { 1, 3 }
        };
        if (cells) {
            foreach (int index in new[] { 0, 3 }) {
                PublisherTableCell cell = table!.Cells[index];
                document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex, Select(cell.Paragraphs));
            }
        } else document.ReplaceStoryParagraphs(story.Id, Select(paragraphs));
        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument saved = PublisherDocument.Load(output);
        if (cells) {
            PublisherTable actual = saved.Pages[page].Tables.Single();
            Assert.Equal(table!.RowHeights, actual.RowHeights); Assert.Equal(table.ColumnWidths, actual.ColumnWidths);
            foreach (int index in new[] { 0, 3 }) AssertParagraphs(order.Select(i => table.Cells[index].Paragraphs[i]).ToArray(), actual.Cells[index].Paragraphs);
            Assert.Equal(table.Cells[1].Text, actual.Cells[1].Text); Assert.Equal(table.Cells[2].Text, actual.Cells[2].Text);
        } else AssertParagraphs(order.Select(i => paragraphs[i]).ToArray(), saved.TextStories[page].Paragraphs);
        Assert.Equal(document.TextStories.Select(s => (s.Id, s.Text)), saved.TextStories.Select(s => (s.Id, s.Text)));
        Assert.Equal(originals.Where(s => s.Id != story.Id).Select(s => (s.Id, s.Text)), saved.TextStories.Where(s => s.Id != story.Id).Select(s => (s.Id, s.Text)));
        PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");

        OfficeRichTextParagraph[] Select(IReadOnlyList<OfficeRichTextParagraph> values) => order.Select(i => operation == "detached"
            ? values[i].WithTabStops(values[i].TabStops) : values[i]).ToArray();
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Copied_empty_paragraphs_with_distinct_opaque_CR_definitions_refuse_both_destination_roles(bool cells, bool finalOnly) {
        PublisherDocument document = PublisherDocument.Load(Fixture(cells));
        PublisherTable? table = cells ? document.Pages[10].Tables.Single() : null;
        var paragraphs = cells ? table!.Cells[0].Paragraphs : document.TextStories[10].Paragraphs;
        Assert.Equal("", paragraphs[0].Runs[0].Text); Assert.Equal("", paragraphs[1].Runs[0].Text);
        OfficeRichTextParagraph copy = paragraphs[0].WithTabStops(paragraphs[0].TabStops);
        var requested = finalOnly ? new[] { copy } : new[] { copy, paragraphs[2] };
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(document.TextStories[10].Id, requested);
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
    }

    private static string Fixture(bool cells) => PublisherNativeTests.Fixture("NativeComparison/"
        + (cells ? "TableNonuniformEmptyParagraphs/native-table" : "TextNonuniformEmptyParagraphs/native-text") + "nonuniformemptyparagraphs.pub");

    private static byte[] Cr(PublisherDocument document, byte[] source, uint storyId, int paragraph) {
        var binding = document.SourceStyleCatalog.Stories.Single(s => s.StoryId == storyId);
        var pp = binding.Uses.Where(u => u.Tag == "FDPP").ElementAt(paragraph);
        int position = pp.NativeStart + pp.NativeLength - 1;
        var span = binding.Uses.Single(u => u.Tag == "FDPC" && u.NativeStart <= position && u.NativeStart + u.NativeLength > position).DirectDefinition;
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        return compound!.Streams["Quill/QuillSub/CONTENTS"].Skip(span.Offset).Take(span.Length).ToArray();
    }

    private static void AssertParagraphs(IReadOnlyList<OfficeRichTextParagraph> expected, IReadOnlyList<OfficeRichTextParagraph> actual) {
        Assert.Equal(expected.Count, actual.Count);
        for (int i = 0; i < expected.Count; i++) {
            var a = expected[i]; var b = actual[i];
            Assert.Equal((a.Alignment, a.Margins, a.Indent, a.LineHeight, a.LineHeightFactor), (b.Alignment, b.Margins, b.Indent, b.LineHeight, b.LineHeightFactor));
            Assert.Equal(a.Runs.Select(r => (r.Text, r.FontFamily, r.FontSize, r.Color, r.Bold, r.Italic)), b.Runs.Select(r => (r.Text, r.FontFamily, r.FontSize, r.Color, r.Bold, r.Italic)));
        }
    }
}
