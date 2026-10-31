using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherFormattingTopologyWritingTests {
    public static IEnumerable<object[]> FormattingCases() {
        foreach (bool cells in new[] { false, true })
            foreach (string operation in new[] { "emphasis", "sizes", "duplicate", "only-final", "layout", "combined",
                "collision", "language-0", "language-1", "equivalent", "first-run", "paragraph-only", "resource-growth" })
                yield return new object[] { cells, operation };
    }

    [Theory]
    [MemberData(nameof(FormattingCases))]
    public void Retained_text_selects_native_templates_for_count_changing_formatting_edits(bool cells, string operation) {
        byte[] source = File.ReadAllBytes(Fixture(cells, operation == "resource-growth"));
        PublisherDocument original = PublisherDocument.Load(source), document = PublisherDocument.Load(source);
        if (operation == "resource-growth") Assert.DoesNotContain("Consolas", original.SourceStyleCatalog.Fonts);
        int page = operation == "equivalent" ? 11 : operation.StartsWith("language", StringComparison.Ordinal) ? 8 : 0;
        PublisherTable? table = cells ? document.Pages[page].Tables.Single() : null;
        var story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[page];
        int[] order = operation switch {
            "sizes" => new[] { 3, 1, 0 }, "duplicate" => new[] { 0, 1, 1, 2, 3 },
            "only-final" => new[] { 3 }, "layout" => new[] { 0 }, "combined" or "resource-growth" => new[] { 2, 0, 3 },
            "collision" or "equivalent" or "language-0" => new[] { 0, page == 0 ? 3 : 2 },
            "language-1" => new[] { 1, 2 }, _ => new[] { 0, 2, 3 }
        };
        OfficeRichTextParagraph[] Select(IReadOnlyList<OfficeRichTextParagraph> values) => order.Select((n, i) => Format(values[n], operation, i)).ToArray();
        if (cells) {
            foreach (int n in new[] { 0, 3 }) {
                var cell = table!.Cells[n]; document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex, Select(cell.Paragraphs));
            }
        } else document.ReplaceStoryParagraphs(story.Id, Select(story.Paragraphs));
        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument saved = PublisherDocument.Load(output);
        if (operation == "resource-growth") {
            Assert.Contains("Consolas", saved.SourceStyleCatalog.Fonts);
            Assert.True(saved.SourceStyleCatalog.Fonts.Count > original.SourceStyleCatalog.Fonts.Count);
            Assert.True(saved.SourceStyleCatalog.Colors.Count > original.SourceStyleCatalog.Colors.Count);
        }
        Assert.Equal(document.TextStories.Select(s => (s.Id, s.Text)), saved.TextStories.Select(s => (s.Id, s.Text)));
        Assert.Equal(original.TextStories.Where(s => s.Id != story.Id).Select(s => (s.Id, s.Text)),
            saved.TextStories.Where(s => s.Id != story.Id).Select(s => (s.Id, s.Text)));
        var beforeParagraphs = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        var actualParagraphs = cells ? saved.Pages[page].Tables.Single().Cells[0].Paragraphs : saved.TextStories[page].Paragraphs;
        var expected = Select(beforeParagraphs);
        Assert.Equal(expected.Length, actualParagraphs.Count);
        for (int i = 0; i < expected.Length; i++) {
            Assert.True(Internal.PublisherTextComparison.ParagraphEqual(expected[i], actualParagraphs[i]));
            bool final = i + 1 == expected.Length;
            int crSource = final ? beforeParagraphs.Count - 1 : order[i];
            bool sourceCr = final || order[i] + 1 < beforeParagraphs.Count;
            Assert.Equal(Definition(original, source, story.Id, crSource, sourceCr ? 0 : 2, sourceCr),
                Definition(saved, output, story.Id, i, 0, cr: true));
            for (int run = 0; run < 3; run++) {
                byte[] before = Definition(original, source, story.Id, order[i], run);
                byte[] after = Definition(saved, output, story.Id, i, run);
                Assert.Equal(Field(before, 0x12), Field(after, 0x12));
                Assert.Equal(Field(before, 0x3E), Field(after, 0x3E));
                if (Internal.PublisherTextComparison.RunFormattingEqual(beforeParagraphs[order[i]].Runs[run], expected[i].Runs[run])) Assert.Equal(before, after);
            }
        }
        if (cells) {
            var actual = saved.Pages[page].Tables.Single();
            Assert.Equal(table!.RowHeights, actual.RowHeights); Assert.Equal(table.ColumnWidths, actual.ColumnWidths);
            Assert.Equal(table.Cells[1].Text, actual.Cells[1].Text); Assert.Equal(table.Cells[2].Text, actual.Cells[2].Text);
            var last = Select(table.Cells[3].Paragraphs);
            Assert.Equal(last.Length, actual.Cells[3].Paragraphs.Count);
            Assert.All(last.Zip(actual.Cells[3].Paragraphs, Internal.PublisherTextComparison.ParagraphEqual), value => Assert.True(value));
        }
        PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");
    }

    [Theory]
    [InlineData(false, false, false)] [InlineData(false, true, false)]
    [InlineData(false, false, true)] [InlineData(false, true, true)]
    [InlineData(true, false, false)] [InlineData(true, true, false)]
    [InlineData(true, false, true)] [InlineData(true, true, true)]
    public void Ambiguous_retained_text_refuses_formatting_even_when_another_source_matches_the_requested_style(bool cells, bool finalOnly, bool styleCollision) {
        byte[] source = File.ReadAllBytes(Fixture(cells)); PublisherDocument document = PublisherDocument.Load(source);
        var table = cells ? document.Pages[8].Tables.Single() : null;
        var story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[8];
        var values = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        Assert.True(Internal.PublisherTextComparison.ParagraphEqual(values[0], values[1]));
        Assert.NotEqual(Field(Definition(document, source, story.Id, 0, 1), 0x12), Field(Definition(document, source, story.Id, 1, 1), 0x12));
        var replacement = styleCollision
            ? Copy(values[0], values[0].Runs, values[2].Alignment, values[2].Margins, values[2].Indent)
            : Format(values[0], "emphasis", 0);
        var requested = finalOnly ? new[] { replacement } : new[] { replacement, values[2] };
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Count_changing_edits_without_retained_text_or_formatting_correspondence_refuse_output(bool cells) {
        PublisherDocument document = PublisherDocument.Load(Fixture(cells));
        var table = cells ? document.Pages[0].Tables.Single() : null;
        var story = cells ? document.TextStories.Single(s => s.Id == table!.StoryId) : document.TextStories[0];
        var values = cells ? table!.Cells[0].Paragraphs : story.Paragraphs;
        var p = Format(values[0], "first-run", 0); var runs = p.Runs.ToArray(); var r = runs[0];
        runs[0] = new OfficeRichTextRun("New text ", r.FontSize, r.Color, r.Bold, r.Italic, r.Underline,
            r.FontFamily, r.Strikethrough, r.BackgroundColor, r.UnderlineStyle, r.StrikethroughStyle, r.Baseline, r.ParagraphIndent);
        var requested = new[] { Copy(p, runs, p.Alignment, p.Margins, p.Indent), values[3] };
        if (cells) document.ReplaceTableCellParagraphs(table!.Id, 0, 0, requested);
        else document.ReplaceStoryParagraphs(story.Id, requested);
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
    }

    private static OfficeRichTextParagraph Format(OfficeRichTextParagraph p, string operation, int index) {
        bool resourceGrowth = operation == "resource-growth";
        if (operation == "resource-growth") operation = "combined";
        bool change = operation.StartsWith("language", StringComparison.Ordinal) ? index == 1
            : operation == "duplicate" ? index is 1 or 2 : index == 0;
        if (!change) return p;
        var runs = p.Runs.Select((r, i) => {
            double size = r.FontSize; bool bold = r.Bold, italic = r.Italic;
            string family = r.FontFamily; OfficeColor color = r.Color; var underline = r.UnderlineStyle;
            if (operation is "emphasis" or "only-final") { if (i == 0) bold = false; if (i == 1) italic = false; }
            if (operation == "sizes") size = new[] { 16D, 10D, 18D }[i];
            if (operation == "duplicate" && index == 1) { if (i == 0) color = OfficeColor.FromRgb(0, 128, 0); if (i == 1) color = OfficeColor.Blue; }
            if (operation == "duplicate" && index == 2 && i == 0) family = "Cambria";
            if (operation == "only-final" && i == 2) { size = 18; bold = true; color = OfficeColor.FromRgb(0, 128, 0); }
            if (operation == "combined" && i == 0) { family = resourceGrowth ? "Consolas" : "Cambria"; size = 16; color = OfficeColor.FromRgb(0, 128, 0); bold = false; underline = OfficeTextDecorationStyle.Double; }
            if (operation == "combined" && i == 1) size = 10;
            if ((operation.StartsWith("language", StringComparison.Ordinal) || operation == "first-run") && i == 0) size = operation == "first-run" ? 18 : 16;
            if (operation == "equivalent" && i == 1) italic = false;
            return new OfficeRichTextRun(r.Text, size, color, bold, italic, false, family, r.Strikethrough, r.BackgroundColor,
                underline, r.StrikethroughStyle, r.Baseline, r.ParagraphIndent) { LinkUri = r.LinkUri, LinkDestinationName = r.LinkDestinationName };
        }).ToArray();
        var alignment = p.Alignment; var margins = p.Margins; var indent = p.Indent;
        if (operation == "layout") { alignment = OfficeTextAlignment.Center; margins = new OfficeTextPadding(12, 4, 0, 9); indent = OfficeTextParagraphIndent.FirstLine(6); }
        if (operation == "combined") { alignment = OfficeTextAlignment.Left; margins = new OfficeTextPadding(margins.Left, margins.Top, margins.Right, 11); }
        if (operation == "collision") { alignment = OfficeTextAlignment.Center; margins = new OfficeTextPadding(0, 2, 0, 6); }
        if (operation == "paragraph-only") margins = new OfficeTextPadding(margins.Left, margins.Top, margins.Right, 11);
        return Copy(p, runs, alignment, margins, indent);
    }

    private static OfficeRichTextParagraph Copy(OfficeRichTextParagraph p, IReadOnlyList<OfficeRichTextRun> runs, OfficeTextAlignment alignment, OfficeTextPadding margins, OfficeTextParagraphIndent indent) =>
        new OfficeRichTextParagraph(runs, alignment, p.LineHeight, margins, indent, p.LineHeightFactor)
            .WithTabStops(new OfficeTextTabStops(p.TabStops!.Stops, p.TabStops.DefaultInterval, -margins.Left)
                .WithParagraphAlignment(alignment is OfficeTextAlignment.Center or OfficeTextAlignment.Right));

    private static string Fixture(bool cells, bool resourceGrowth = false) => PublisherNativeTests.Fixture("NativeComparison/"
        + (resourceGrowth ? (cells ? "TableMixedRunTextReplacements/native-table" : "TextMixedRunTextReplacements/native-text") + "mixedruntextreplacements.pub"
            : (cells ? "TableFormattingTopology/native-table" : "TextFormattingTopology/native-text") + "formattingtopology.pub"));

    private static byte[] Definition(PublisherDocument document, byte[] source, uint id, int paragraph, int run, bool cr = false) {
        var binding = document.SourceStyleCatalog.Stories.Single(s => s.StoryId == id); var pp = binding.Uses.Where(u => u.Tag == "FDPP").ElementAt(paragraph);
        int position = pp.NativeStart + (cr ? pp.NativeLength - 1 : document.TextStories.Single(s => s.Id == id).Paragraphs[paragraph].Runs.Take(run).Sum(r => r.Text.Length));
        var span = binding.Uses.Single(u => u.Tag == "FDPC" && u.NativeStart <= position && u.NativeStart + u.NativeLength > position).DirectDefinition;
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        return compound!.Streams["Quill/QuillSub/CONTENTS"].Skip(span.Offset).Take(span.Length).ToArray();
    }

    private static byte[]? Field(byte[] definition, byte id) {
        var data = new Internal.PublisherBinaryData(definition, "test/native-definition");
        var context = new Internal.PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        var fields = new Internal.PublisherBlockReader(data, context).Chunk(0, definition.Length).Where(f => f.Id == id).ToArray();
        if (fields.Length == 0) return null;
        var field = Assert.Single(fields);
        return definition.Skip(field.Offset - 2).Take(field.Length + 2).ToArray();
    }
}
