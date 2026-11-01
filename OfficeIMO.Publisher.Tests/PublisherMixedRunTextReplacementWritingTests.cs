using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherMixedRunTextReplacementWritingTests {
    public static IEnumerable<object[]> ReplacementCases() {
        foreach (bool cells in new[] { false, true })
            foreach (string operation in new[] { "grow", "shrink", "reorder", "duplicate", "only-final", "only-first",
                "equivalent", "middle-only", "language-0", "language-1" })
                yield return new object[] { cells, operation };
    }

    [Theory]
    [MemberData(nameof(ReplacementCases))]
    public void Mixed_run_text_replacements_preserve_selected_native_templates_during_count_changes(bool cells, string operation) {
        byte[] source = File.ReadAllBytes(Fixture(cells));
        PublisherDocument document = PublisherDocument.Load(source);
        var originals = document.TextStories.ToArray();
        PublisherDocument nativeSource = PublisherDocument.Load(source);
        int page = operation == "equivalent" ? 8 : operation.StartsWith("language", StringComparison.Ordinal) ? 7 : 0;
        PublisherTable? table = cells ? document.Pages[page].Tables.Single() : null;
        PublisherTextStory story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[page];
        var original = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        int[] order = operation switch {
            "grow" => new[] { 0, 2, 3 }, "shrink" => new[] { 1, 2, 3 }, "reorder" => new[] { 3, 1, 0 },
            "duplicate" => new[] { 0, 1, 1, 2, 3 }, "only-final" => new[] { 3 }, "only-first" => new[] { 0 },
            "equivalent" => new[] { 0, 2 }, "middle-only" => new[] { 0, 2, 3 }, "language-0" => new[] { 0, 2 }, _ => new[] { 1, 2 }
        };
        string?[] text = operation switch {
            "grow" => new[] { "Expanded lengthened ending", "Updated rewritten result", null },
            "shrink" => new[] { "S b t", "T c z", null },
            "reorder" => new[] { "Finale changed ending", "Second altered tail", "First edited finish" },
            "duplicate" => new[] { null, "First centered edit", "Second centered edit", null, null },
            "only-final" => new[] { "Finale modified ending" }, "only-first" => new[] { "Ωmega replaced finish" },
            "equivalent" => new[] { "Common edited result", null }, "middle-only" => new[] { "First altered tail", null, null },
            _ => new[] { null, "Finale modified ending" }
        };
        OfficeRichTextParagraph[] Select(IReadOnlyList<OfficeRichTextParagraph> values) => order.Select((index, i) =>
            text[i] is string replacement ? ReplaceText(values[index], replacement) : values[index]).ToArray();
        var requested = Select(original);
        if (cells) {
            foreach (int index in new[] { 0, 3 }) {
                var cell = table!.Cells[index];
                document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex, Select(cell.Paragraphs));
            }
        } else document.ReplaceStoryParagraphs(story.Id, requested);

        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument saved = PublisherDocument.Load(output);
        Assert.Equal(document.TextStories.Select(s => (s.Id, s.Text)), saved.TextStories.Select(s => (s.Id, s.Text)));
        Assert.Equal(originals.Where(s => s.Id != story.Id).Select(s => (s.Id, s.Text)),
            saved.TextStories.Where(s => s.Id != story.Id).Select(s => (s.Id, s.Text)));
        if (cells) {
            var actual = saved.Pages[page].Tables.Single();
            Assert.Equal(table!.ColumnWidths, actual.ColumnWidths); Assert.Equal(table.RowHeights, actual.RowHeights);
            Assert.Equal(table.Cells[1].Text, actual.Cells[1].Text); Assert.Equal(table.Cells[2].Text, actual.Cells[2].Text);
            foreach (int index in new[] { 0, 3 }) Check(Select(table.Cells[index].Paragraphs), actual.Cells[index].Paragraphs,
                index == 0 ? 0 : original.Count + 2, index == 0 ? 0 : requested.Length + 2);
        } else Check(requested, saved.TextStories[page].Paragraphs, 0, 0);
        PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");

        void Check(IReadOnlyList<OfficeRichTextParagraph> expected, IReadOnlyList<OfficeRichTextParagraph> actual, int beforeOffset, int afterOffset) {
            Assert.Equal(expected.Count, actual.Count);
            for (int i = 0; i < expected.Count; i++) {
                Assert.Equal(string.Concat(expected[i].Runs.Select(r => r.Text)), string.Concat(actual[i].Runs.Select(r => r.Text)));
                Assert.True(Internal.PublisherTextComparison.ParagraphEqual(expected[i], actual[i]));
                Assert.Equal(Definition(nativeSource, source, story.Id, beforeOffset + order[i], "FDPP"),
                    Definition(saved, output, story.Id, afterOffset + i, "FDPP"));
                Assert.Equal(3, expected[i].Runs.Count);
                for (int run = 0; run < 3; run++)
                    Assert.Equal(Definition(nativeSource, source, story.Id, beforeOffset + order[i], "FDPC", run: run),
                        Definition(saved, output, story.Id, afterOffset + i, "FDPC", run: run));
                int crSource = expected[i].Runs[0].Text.Length == 0 ? order[i]
                    : i + 1 == expected.Count ? original.Count - 1 : order[i];
                bool useCr = expected[i].Runs[0].Text.Length == 0 || i + 1 == expected.Count || order[i] + 1 < original.Count;
                Assert.Equal(Definition(nativeSource, source, story.Id, beforeOffset + crSource, "FDPC", useCr, useCr ? 0 : 2),
                    Definition(saved, output, story.Id, afterOffset + i, "FDPC", true));
            }
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Changed_text_refuses_indistinguishable_formatting_with_conflicting_native_language_templates(bool cells, bool finalOnly) {
        byte[] source = File.ReadAllBytes(Fixture(cells));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable? table = cells ? document.Pages[7].Tables.Single() : null;
        var story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[7];
        var values = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        Assert.True(Internal.PublisherTextComparison.ParagraphEqual(values[0], values[1]));
        Assert.NotEqual(Definition(document, source, story.Id, 0, "FDPC", run: 1), Definition(document, source, story.Id, 1, "FDPC", run: 1));
        var replacement = ReplaceText(values[0], "Ambiguous middle replacement");
        var requested = finalOnly ? new[] { replacement } : new[] { replacement, values[2] };
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Mixed_run_count_changes_refuse_empty_replacement_runs(bool cells) {
        PublisherDocument document = PublisherDocument.Load(Fixture(cells));
        PublisherTable? table = cells ? document.Pages[0].Tables.Single() : null;
        var story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[0];
        var values = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        OfficeRichTextParagraph p = ReplaceText(values[0], "Changed styled paragraph");
        var runs = p.Runs.ToList();
        OfficeRichTextRun Copy(OfficeRichTextRun r, string text) => new OfficeRichTextRun(text, r.FontSize, r.Color,
            r.Bold, r.Italic, r.Underline, r.FontFamily, r.Strikethrough, r.BackgroundColor,
            r.UnderlineStyle, r.StrikethroughStyle, r.Baseline, r.ParagraphIndent);
        runs[1] = Copy(runs[1], "");
        var replacement = new OfficeRichTextParagraph(runs, p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
        var requested = new[] { replacement, values[3] };
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
    }

    private static OfficeRichTextParagraph ReplaceText(OfficeRichTextParagraph p, string text) {
        string[] words = text.Split(' ');
        Assert.Equal(3, words.Length); Assert.Equal(3, p.Runs.Count);
        return new OfficeRichTextParagraph(p.Runs.Select((r, i) => new OfficeRichTextRun(words[i] + (i < 2 ? " " : ""),
            r.FontSize, r.Color, r.Bold, r.Italic, r.Underline, r.FontFamily, r.Strikethrough,
            r.BackgroundColor, r.UnderlineStyle, r.StrikethroughStyle, r.Baseline, r.ParagraphIndent) {
                LinkUri = r.LinkUri, LinkDestinationName = r.LinkDestinationName
            }).ToArray(), p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
    }

    private static string Fixture(bool cells) => PublisherNativeTests.Fixture("NativeComparison/"
        + (cells ? "TableMixedRunTextReplacements/native-table" : "TextMixedRunTextReplacements/native-text") + "mixedruntextreplacements.pub");

    private static byte[] Definition(PublisherDocument document, byte[] source, uint id, int paragraph, string tag, bool cr = false, int run = 0) {
        var binding = document.SourceStyleCatalog.Stories.Single(s => s.StoryId == id);
        var p = binding.Uses.Where(u => u.Tag == "FDPP").ElementAt(paragraph);
        var paragraphValue = document.TextStories.Single(s => s.Id == id).Paragraphs[paragraph];
        int position = p.NativeStart + (cr ? p.NativeLength - 1 : paragraphValue.Runs.Take(run).Sum(r => r.Text.Length));
        var span = binding.Uses.Single(u => u.Tag == tag && u.NativeStart <= position && u.NativeStart + u.NativeLength > position).DirectDefinition;
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        return compound!.Streams["Quill/QuillSub/CONTENTS"].Skip(span.Offset).Take(span.Length).ToArray();
    }
}
