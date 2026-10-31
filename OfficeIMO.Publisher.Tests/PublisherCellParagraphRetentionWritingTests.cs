using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherCellParagraphRetentionWritingTests {
    private const string Fixture = "NativeComparison/TableNonuniformParagraphs/native-tablenonuniformparagraphs.pub";

    [Theory]
    [InlineData("drop-first")]
    [InlineData("drop-middle")]
    [InlineData("drop-final")]
    [InlineData("reorder")]
    [InlineData("duplicate")]
    [InlineData("detached")]
    [InlineData("language-0")]
    [InlineData("language-1")]
    [InlineData("first-only")]
    [InlineData("final-only")]
    public void Nonuniform_cell_retention_preserves_partitions_styles_and_neighboring_native_storage(string operation) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(Fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable[] originals = document.Pages.SelectMany(p => p.Tables).ToArray();
        PublisherTable table = originals[operation.StartsWith("language", StringComparison.Ordinal) ? 6 : 0];
        int[] order = operation switch {
            "drop-first" or "final-only" => new[] { 1, 2, 3 }, "drop-middle" => new[] { 0, 2, 3 },
            "drop-final" or "first-only" => new[] { 0, 1, 2 }, "reorder" or "detached" => new[] { 3, 0, 2 },
            "duplicate" => new[] { 0, 1, 1, 2, 3 }, "language-0" => new[] { 0, 2 }, _ => new[] { 1, 2 }
        };
        int[] targets = operation == "first-only" ? new[] { 0 } : operation == "final-only" ? new[] { 3 } : new[] { 0, 3 };
        foreach (int index in targets) {
            PublisherTableCell cell = table.Cells[index];
            var selected = order.Select(i => operation == "detached" ? cell.Paragraphs[i].WithTabStops(cell.Paragraphs[i].TabStops)
                : cell.Paragraphs[i]).ToArray();
            document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex, selected);
        }
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        byte[] bytes = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument saved = PublisherDocument.Load(bytes);
        foreach (PublisherTable before in originals) {
            PublisherTable after = saved.Pages.SelectMany(p => p.Tables).Single(t => t.Id == before.Id);
            Assert.Equal((before.StoryId, before.PageId, before.X, before.Y, before.Width, before.Height),
                (after.StoryId, after.PageId, after.X, after.Y, after.Width, after.Height));
            Assert.Equal(before.RowHeights, after.RowHeights); Assert.Equal(before.ColumnWidths, after.ColumnWidths);
            for (int i = 0; i < before.Cells.Count; i++) {
                PublisherTableCell a = before.Cells[i], b = after.Cells[i];
                Assert.Equal((a.RowIndex, a.ColumnIndex, a.RowSpan, a.ColumnSpan, a.X, a.Y, a.Width, a.Height),
                    (b.RowIndex, b.ColumnIndex, b.RowSpan, b.ColumnSpan, b.X, b.Y, b.Width, b.Height));
                var expected = before.Id == table.Id && targets.Contains(i) ? order.Select(j => a.Paragraphs[j]).ToArray() : a.Paragraphs.ToArray();
                Assert.Equal(string.Join("\n", expected.Select(p => string.Concat(p.Runs.Select(r => r.Text)))), b.Text);
                Assert.Equal(expected.Length, b.Paragraphs.Count);
                for (int p = 0; p < expected.Length; p++) {
                    OfficeRichTextParagraph x = expected[p], y = b.Paragraphs[p];
                    Assert.Equal((x.Alignment, x.Margins, x.Indent, x.LineHeight, x.LineHeightFactor),
                        (y.Alignment, y.Margins, y.Indent, y.LineHeight, y.LineHeightFactor));
                    Assert.Equal(x.Runs.Select(r => (r.Text, r.FontFamily, r.FontSize, r.Bold, r.Italic, r.Color)),
                        y.Runs.Select(r => (r.Text, r.FontFamily, r.FontSize, r.Bold, r.Italic, r.Color)));
                }
            }
        }
        Assert.Equal(document.TextStories.Select(s => (s.Id, s.Text)), saved.TextStories.Select(s => (s.Id, s.Text)));
        Assert.Equal(operation.StartsWith("language", StringComparison.Ordinal) ? 3 : 4, table.Cells[0].Paragraphs.Count);
        PublisherNativeWritingTests.AssertPreservedStorage(source, bytes, "Quill/QuillSub/CONTENTS");
    }

    [Theory]
    [InlineData(0)]
    [InlineData(3)]
    public void Copied_cell_paragraphs_refuse_ambiguous_native_language_definitions(int cellIndex) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture(Fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable table = document.Pages[6].Tables.Single(); PublisherTableCell cell = table.Cells[cellIndex];
        Assert.True(table.CanEditCellText); Assert.Equal(source, document.ToPubBytes());
        Assert.Equal(cell.Paragraphs[0].Runs[0].Text, cell.Paragraphs[1].Runs[0].Text);
        OfficeRichTextParagraph paragraph = cell.Paragraphs[0];
        document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex,
            new[] { paragraph.WithTabStops(paragraph.TabStops), cell.Paragraphs[2] });
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.True(Assert.IsType<PublisherSaveReport>(error.Report).HasErrors);
        Assert.Equal(3, cell.Paragraphs.Count);
    }
}
