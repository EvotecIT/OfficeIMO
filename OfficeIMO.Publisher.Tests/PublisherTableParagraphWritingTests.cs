using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableParagraphWritingTests {
    [Theory]
    [InlineData("first")]
    [InlineData("final")]
    [InlineData("all")]
    public void Paragraph_growth_retains_each_cells_text_style_and_table_ownership(string edit) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableLabel/native-tablelabel.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = Tables(document).Single();
        int[] targets = edit == "all" ? new[] { 0, 1, 2, 3 } : new[] { edit == "first" ? 0 : 3 };
        foreach (int index in targets) {
            PublisherTableCell cell = table.Cells[index];
            OfficeRichTextParagraph p = cell.Paragraphs[0];
            document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex,
                new[] { Paragraph(p, "Cell " + index + " first"), Paragraph(p, "Cell " + index + " second") });
        }
        PublisherDocument saved = SaveAndCheck(document, table, source);
        PublisherTable actual = Tables(saved).Single();
        foreach (int index in Enumerable.Range(0, 4)) {
            PublisherTableCell cell = actual.Cells[index];
            if (targets.Contains(index)) {
                Assert.Equal(new[] { "Cell " + index + " first", "Cell " + index + " second" },
                    cell.Paragraphs.Select(p => string.Concat(p.Runs.Select(run => run.Text))));
            } else Assert.Equal(table.Cells[index].Text, cell.Text);
            Assert.All(cell.Paragraphs, p => Assert.Equal(index < 2, p.Runs[0].Bold));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Native_multiparagraph_cells_can_shrink_without_changing_neighbors_or_table_grid(bool balanced) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableCellParagraphs/native-tablecellparagraphs.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = Tables(document).Last();
        Assert.All(table.Cells, cell => Assert.Equal(2, cell.Paragraphs.Count));
        PublisherTableCell final = table.Cells[3];
        document.ReplaceTableCellParagraphs(table.Id, final.RowIndex, final.ColumnIndex, new[] { Paragraph(final.Paragraphs[0], "Last shrank") });
        if (balanced) {
            PublisherTableCell first = table.Cells[0];
            document.ReplaceTableCellParagraphs(table.Id, first.RowIndex, first.ColumnIndex,
                new[] { Paragraph(first.Paragraphs[0], "First a"), Paragraph(first.Paragraphs[0], "First b"), Paragraph(first.Paragraphs[0], "First c") });
        }
        PublisherDocument saved = SaveAndCheck(document, table, source);
        PublisherTable actual = Tables(saved).Single(t => t.Id == table.Id);
        Assert.Equal(balanced ? new[] { 3, 2, 2, 1 } : new[] { 2, 2, 2, 1 }, actual.Cells.Select(c => c.Paragraphs.Count));
        Assert.Equal("Last shrank", actual.Cells[3].Text);
        Assert.All(actual.Cells[0].Paragraphs, p => Assert.True(p.Runs[0].Bold));
        Assert.False(actual.Cells[3].Paragraphs[0].Runs[0].Bold);
    }

    [Fact]
    public void Moving_a_paragraph_between_cells_remaps_native_partitions_when_the_complete_styled_story_is_unchanged() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableCellParagraphs/native-tablecellparagraphs.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = Tables(document).Last();
        PublisherTextStory before = document.TextStories.Single(s => s.Id == table.StoryId);
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] { table.Cells[0].Paragraphs[0] });
        document.ReplaceTableCellParagraphs(table.Id, 0, 1, new[] { table.Cells[0].Paragraphs[1] }.Concat(table.Cells[1].Paragraphs).ToArray());
        PublisherTextStory after = document.TextStories.Single(s => s.Id == table.StoryId);
        Assert.Equal(before.Text, after.Text);
        Assert.Equal(before.Paragraphs.Select(p => (p.Alignment, p.LineHeight, p.LineHeightFactor)),
            after.Paragraphs.Select(p => (p.Alignment, p.LineHeight, p.LineHeightFactor)));
        Assert.Equal(before.Paragraphs.SelectMany(p => p.Runs).Select(r => (r.Text, r.FontFamily, r.FontSize, r.Bold, r.Italic, r.Color)),
            after.Paragraphs.SelectMany(p => p.Runs).Select(r => (r.Text, r.FontFamily, r.FontSize, r.Bold, r.Italic, r.Color)));
        PublisherTable saved = Tables(SaveAndCheck(document, table, source)).Single(t => t.Id == table.Id);
        Assert.Equal(new[] { 1, 3, 2, 2 }, saved.Cells.Select(c => c.Paragraphs.Count));
        Assert.Equal("Item one", saved.Cells[0].Text);
        Assert.Equal("Item two\nCount one\nCount two", saved.Cells[1].Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Appended_empty_cell_paragraphs_retain_or_explicitly_set_the_final_CR_typing(bool explicitTyping) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableLabel/native-tablelabel.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = Tables(document).Single();
        OfficeRichTextParagraph p = table.Cells[0].Paragraphs[0];
        OfficeRichTextRun[] runs = explicitTyping
            ? new[] { new OfficeRichTextRun("", 14, OfficeColor.Blue, bold: true, italic: true, fontFamily: "Times New Roman") }
            : Array.Empty<OfficeRichTextRun>();
        var empty = new OfficeRichTextParagraph(runs, p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] { p, empty });
        PublisherTable saved = Tables(SaveAndCheck(document, table, source)).Single();
        OfficeRichTextRun actual = Assert.Single(saved.Cells[0].Paragraphs[1].Runs);
        Assert.Equal("", actual.Text);
        if (explicitTyping) Assert.Equal(("Times New Roman", 14D, true, true, OfficeColor.Blue),
            (actual.FontFamily, actual.FontSize, actual.Bold, actual.Italic, actual.Color));
        else Assert.Equal(("Arial", 10D, true, false, OfficeColor.Black),
            (actual.FontFamily, actual.FontSize, actual.Bold, actual.Italic, actual.Color));
    }

    [Fact]
    public void Nonuniform_definitions_within_the_changed_cell_still_refuse_paragraph_count_changes() {
        PublisherDocument original = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/TableCellParagraphs/native-tablecellparagraphs.pub"));
        PublisherTable table = Tables(original).Last();
        OfficeRichTextParagraph[] paragraphs = table.Cells[0].Paragraphs.ToArray();
        OfficeRichTextParagraph p = paragraphs[1];
        paragraphs[1] = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Different", 14, OfficeColor.Blue,
            bold: true, italic: true, fontFamily: "Times New Roman") }, p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
        original.ReplaceTableCellParagraphs(table.Id, 0, 0, paragraphs);
        PublisherDocument source = PublisherDocument.Load(original.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        PublisherTable reloaded = Tables(source).Single(t => t.Id == table.Id);
        Assert.Equal(14D, reloaded.Cells[0].Paragraphs[1].Runs[0].FontSize);
        source.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] { Paragraph(reloaded.Cells[0].Paragraphs[0], "Retained") });
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => source.ToPubBytes(
            new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.LossKind == OfficeConversionLossKind.Failure);
    }

    private static PublisherDocument SaveAndCheck(PublisherDocument document, PublisherTable original, byte[] source) {
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        byte[] bytes = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument saved = PublisherDocument.Load(bytes);
        PublisherTable expected = Tables(document).Single(t => t.Id == original.Id), actual = Tables(saved).Single(t => t.Id == original.Id);
        Assert.Equal((original.Id, original.StoryId, original.PageId, original.X, original.Y, original.Width, original.Height),
            (actual.Id, actual.StoryId, actual.PageId, actual.X, actual.Y, actual.Width, actual.Height));
        Assert.Equal(original.ColumnWidths, actual.ColumnWidths); Assert.Equal(original.RowHeights, actual.RowHeights);
        Assert.Equal(expected.Cells.Select(c => c.Text), actual.Cells.Select(c => c.Text));
        Assert.Equal(expected.Cells.Select(c => c.Paragraphs.Count), actual.Cells.Select(c => c.Paragraphs.Count));
        Assert.Equal(original.Cells.Select(c => (c.RowIndex, c.ColumnIndex, c.RowSpan, c.ColumnSpan)),
            actual.Cells.Select(c => (c.RowIndex, c.ColumnIndex, c.RowSpan, c.ColumnSpan)));
        Assert.Equal(document.TextStories.Select(s => (s.Id, s.Text)), saved.TextStories.Select(s => (s.Id, s.Text)));
        PublisherNativeWritingTests.AssertPreservedStorage(source, bytes, "Quill/QuillSub/CONTENTS");
        return saved;
    }

    private static IEnumerable<PublisherTable> Tables(PublisherDocument document) =>
        document.Pages.Concat(document.MasterPages).SelectMany(page => page.Tables);

    private static OfficeRichTextParagraph Paragraph(OfficeRichTextParagraph p, string text) =>
        new OfficeRichTextParagraph(p.Runs.Select((run, i) => new OfficeRichTextRun(i == 0 ? text : run.Text,
            run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough,
            run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle, run.Baseline, run.ParagraphIndent)).ToArray(),
            p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
}
