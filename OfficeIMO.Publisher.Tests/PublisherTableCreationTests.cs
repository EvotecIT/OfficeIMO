using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableCreationTests {
    [Theory]
    [InlineData("TableCreation/native-tablecreation.pub", false)]
    [InlineData("GeneratedTable/generated-table.pub", false)]
    [InlineData("GeneratedTable/native-normalized.pub", false)]
    [InlineData("GeneratedTable/generated-edited.pub", true)]
    public void Native_created_table_keeps_cell_boundaries_tracks_and_opaque_layout_across_growth_and_clear_saves(string fixture, bool initiallyEdited) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = Assert.Single(document.Pages.Single().Tables);
        Check(document, initiallyEdited ? new[] { "Updated\nItem", "Count", "Alpha", "7" } : new[] { "Item", "Count", "Alpha", "" });

        foreach (bool clear in new[] { initiallyEdited, !initiallyEdited }) {
            PublisherTable table = document.Pages.Single().Tables.Single();
            foreach (int index in new[] { 0, 3 }) {
                PublisherTableCell cell = table.Cells[index];
                OfficeRichTextParagraph template = cell.Paragraphs[0];
                OfficeRichTextRun run = template.Runs[0];
                OfficeRichTextParagraph[] paragraphs = clear ? Array.Empty<OfficeRichTextParagraph>()
                    : (index == 0 ? new[] { "Updated", "Item" } : new[] { "7" }).Select(text =>
                        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, run.FontSize, run.Color,
                            run.Bold, run.Italic, run.Underline, run.FontFamily) }, template.Alignment, template.LineHeight,
                            template.Margins, template.Indent, template.LineHeightFactor).WithTabStops(template.TabStops)).ToArray();
                document.ReplaceTableCellParagraphs(table.Id, cell.RowIndex, cell.ColumnIndex, paragraphs);
            }
            var saved = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            Assert.False(saved.Report.HasErrors);
            PublisherNativeWritingTests.AssertPreservedStorage(source, saved.Value, "Quill/QuillSub/CONTENTS");
            Assert.Equal(Layout(source), Layout(saved.Value));
            document = PublisherDocument.Load(saved.Value);
            PublisherTable reopened = document.Pages.Single().Tables.Single();
            Assert.Equal((original.Id, original.StoryId, original.PageId), (reopened.Id, reopened.StoryId, reopened.PageId));
            Check(document, clear ? new[] { "", "Count", "Alpha", "" } : new[] { "Updated\nItem", "Count", "Alpha", "7" });
            source = saved.Value;
        }
    }

    private static void Check(PublisherDocument document, string[] text) {
        PublisherTable table = Assert.Single(document.Pages.Single().Tables);
        Assert.True(table.HasTextMapping);
        Assert.True(table.CanEditCellText);
        Assert.Equal((72D, 144D, 300D, 96D), (table.X, table.Y, table.Width, table.Height));
        Assert.Equal(new[] { 120D, 180D }, table.ColumnWidths);
        Assert.Equal(new[] { 42D, 54D }, table.RowHeights);
        Assert.Equal(new[] { (0, 0), (0, 1), (1, 0), (1, 1) }, table.Cells.Select(cell => (cell.RowIndex, cell.ColumnIndex)));
        Assert.All(table.Cells, cell => Assert.Equal((1, 1), (cell.RowSpan, cell.ColumnSpan)));
        Assert.Equal(text, table.Cells.Select(cell => cell.Text));
        Assert.Equal(2, document.TextStories.Count);
        Assert.Equal(string.Join("\n", text) + "\n", document.TextStories.Single(story => story.Id == table.StoryId).Text);
        Assert.Equal("Independent table label\n", Assert.Single(document.TextStories, story => story.Id != table.StoryId).Text);
        foreach (PublisherTableCell cell in table.Cells) {
            Assert.All(cell.Paragraphs, paragraph => {
                Assert.Equal(18D, paragraph.LineHeight);
                Assert.All(paragraph.Runs, run => { Assert.Equal(12D, run.FontSize); Assert.Equal(cell.RowIndex == 0, run.Bold); });
            });
        }
    }

    private static byte[] Layout(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] stream = compound!.Streams["Quill/QuillSub/CONTENTS"];
        var data = new PublisherBinaryData(stream, "Quill/QuillSub/CONTENTS");
        PublisherQuillChunk chunk = PublisherQuillDirectory.Read(data,
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None)).Single(value => value.Tag == "MCLD");
        return stream.Skip(chunk.Offset).Take(chunk.Length).ToArray();
    }
}
