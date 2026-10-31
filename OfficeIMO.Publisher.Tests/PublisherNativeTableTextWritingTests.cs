using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherNativeTableTextWritingTests {
    [Theory]
    [InlineData("first")]
    [InlineData("final")]
    [InlineData("all")]
    [InlineData("ordinary")]
    public void Native_saves_retain_table_cell_ownership_when_cell_or_preceding_story_lengths_change(string edit) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableLabel/native-tablelabel.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = document.Pages[0].Tables.Single();
        if (edit == "ordinary") {
            PublisherTextStory title = document.TextStories.Single(story => story.Id != original.StoryId);
            document.ReplaceStoryParagraphs(title.Id, Replace(title.Paragraphs[0], "A longer title before the table story."));
        } else {
            int[] indices = edit == "all" ? new[] { 0, 1, 2, 3 } : new[] { edit == "first" ? 0 : 3 };
            foreach (int index in indices) {
                PublisherTableCell cell = original.Cells[index];
                document.ReplaceTableCellParagraphs(original.Id, cell.RowIndex, cell.ColumnIndex,
                    Replace(cell.Paragraphs[0], index == 0 ? "Expanded item" : "Value " + index, edit == "all"));
            }
        }
        PublisherTable expected = document.Pages[0].Tables.Single();
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        byte[] output = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument reopened = PublisherDocument.Load(output);
        PublisherTable actual = reopened.Pages[0].Tables.Single();
        Assert.Equal((original.Id, original.StoryId, original.PageId, original.X, original.Y, original.Width, original.Height),
            (actual.Id, actual.StoryId, actual.PageId, actual.X, actual.Y, actual.Width, actual.Height));
        Assert.Equal(original.ColumnWidths, actual.ColumnWidths); Assert.Equal(original.RowHeights, actual.RowHeights);
        Assert.Equal(expected.Cells.Select(cell => cell.Text), actual.Cells.Select(cell => cell.Text));
        Assert.Equal(expected.Cells.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan)),
            actual.Cells.Select(cell => (cell.RowIndex, cell.ColumnIndex, cell.RowSpan, cell.ColumnSpan)));
        Assert.Equal(document.TextStories.Select(story => (story.Id, story.Text)), reopened.TextStories.Select(story => (story.Id, story.Text)));
        if (edit == "all") Assert.All(actual.Cells, cell => Assert.True(cell.Paragraphs[0].Runs[0].Italic));
        else Assert.Equal(original.Cells.Select(cell => cell.Paragraphs[0].Runs[0].Bold), actual.Cells.Select(cell => cell.Paragraphs[0].Runs[0].Bold));
        PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS");
        Assert.Equal(new[] { "Item", "Count", "Alpha", "2" }, original.Cells.Select(cell => cell.Text));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unknown_native_cell_partition_metadata_refuses_edited_output_without_weakening_unchanged_copying(bool descriptor) {
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
            PublisherQuillChunk chunk = PublisherQuillDirectory.Read(new PublisherBinaryData(bytes, "Quill"), context)
                .Single(item => item.Tag == "TCD ");
            if (descriptor) {
                PublisherQuillSourceLayout layout = PublisherQuillSourceLayout.TryRead(bytes, context)!;
                PublisherQuillChunk bound = layout.Chunks.Single(item => item.Tag == "TCD ");
                bytes[layout.DescriptorOffsets[bound] + 8] = 2;
            } else bytes[chunk.Offset + 8] = 1;
        }, "NativeComparison/TableLabel/native-tablelabel.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        PublisherTable table = document.Pages[0].Tables.Single();
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, Replace(table.Cells[0].Paragraphs[0], "Changed item"));
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(
            new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.LossKind == OfficeConversionLossKind.Failure);
        Assert.Equal("Changed item", document.Pages[0].Tables[0].Cells[0].Text);
    }

    private static OfficeRichTextParagraph[] Replace(OfficeRichTextParagraph template, string text, bool italic = false) {
        return new[] { new OfficeRichTextParagraph(template.Runs.Select((run, index) =>
            new OfficeRichTextRun(index == 0 ? text : run.Text, run.FontSize, run.Color, run.Bold, italic || run.Italic, run.Underline, run.FontFamily,
                run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle, run.Baseline, run.ParagraphIndent)
        ).ToArray(), template.Alignment, template.LineHeight, template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops) };
    }
}
