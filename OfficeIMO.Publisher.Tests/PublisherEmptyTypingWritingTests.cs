using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherEmptyTypingWritingTests {
    [Theory]
    [InlineData("TableLabel", false, 0)]
    [InlineData("TableLabel", false, 3)]
    [InlineData("TableLabel", true, 0)]
    [InlineData("TableCellTyping", true, 0)]
    [InlineData("TableCellTypingEarlyRange", false, 0)]
    [InlineData("TableCellTypingEarlyRange", true, 0)]
    public void Cleared_cells_save_their_native_CR_or_the_explicitly_requested_empty_typing_style(string fixture, bool explicitStyle, int cellIndex) {
        string leaf = fixture == "TableLabel" ? "native-tablelabel.pub" : "native-tablecelltyping.pub";
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture($"NativeComparison/{fixture}/{leaf}"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTable original = Tables(document).First();
        var report = document.SourceReadReport; var diagnostics = report.FidelityDiagnostics.ToArray();
        PublisherTableCell target = original.Cells[cellIndex];
        if (explicitStyle) {
            OfficeRichTextParagraph p = target.Paragraphs[0];
            document.ReplaceTableCellParagraphs(original.Id, target.RowIndex, target.ColumnIndex, new[] {
                new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("", 14, OfficeColor.Blue, bold: true, italic: true, fontFamily: "Times New Roman") },
                    p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops)
            });
        } else {
            document.ReplaceTableCellParagraphs(original.Id, target.RowIndex, target.ColumnIndex, Array.Empty<OfficeRichTextParagraph>());
            OfficeRichTextParagraph cleared = Tables(document).Single(table => table.Id == original.Id).Cells[cellIndex].Paragraphs[0];
            Assert.Equal(target.Paragraphs.Last().LineHeightFactor, cleared.LineHeightFactor);
            Assert.Equal(target.Paragraphs.Last().Margins, cleared.Margins);
        }
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        byte[] bytes = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherTable saved = Tables(PublisherDocument.Load(bytes)).Single(table => table.Id == original.Id);
        Assert.Equal("", saved.Cells[cellIndex].Text);
        Assert.Single(saved.Cells[cellIndex].Paragraphs);
        OfficeRichTextRun run = Assert.Single(saved.Cells[cellIndex].Paragraphs[0].Runs);
        if (explicitStyle) {
            Assert.Equal(("Times New Roman", 14D, true, true, OfficeColor.Blue), (run.FontFamily, run.FontSize, run.Bold, run.Italic, run.Color));
        } else {
            string control = fixture == "TableCellTypingEarlyRange" ? fixture : "TableCellTyping";
            PublisherTable[] native = Tables(PublisherDocument.Load(PublisherNativeTests.Fixture($"NativeComparison/{control}/native-tablecelltyping.pub"))).ToArray();
            OfficeRichTextRun expected = native[cellIndex == 3 ? 2 : 1].Cells[cellIndex].Paragraphs[0].Runs[0];
            Assert.Equal((expected.FontFamily, expected.FontSize, expected.Bold, expected.Italic, expected.Color),
                (run.FontFamily, run.FontSize, run.Bold, run.Italic, run.Color));
        }
        Assert.Equal(original.ColumnWidths, saved.ColumnWidths); Assert.Equal(original.RowHeights, saved.RowHeights);
        Assert.Equal((original.Id, original.StoryId, original.PageId, original.X, original.Y, original.Width, original.Height),
            (saved.Id, saved.StoryId, saved.PageId, saved.X, saved.Y, saved.Width, saved.Height));
        for (int i = 0; i < original.Cells.Count; i++) if (i != cellIndex) Assert.Equal(original.Cells[i].Text, saved.Cells[i].Text);
        PublisherNativeWritingTests.AssertPreservedStorage(source, bytes, "Quill/QuillSub/CONTENTS");
        Assert.Same(report, document.SourceReadReport); Assert.Equal(diagnostics, report.FidelityDiagnostics);
        Assert.Equal("Item", original.Cells[0].Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void A_retained_single_paragraph_can_be_empty_with_explicit_or_unspecified_typing(bool explicitStyle) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"));
        PublisherTextStory original = document.TextStories.Single(); OfficeRichTextParagraph p = original.Paragraphs[0];
        document.ReplaceStoryParagraphs(original.Id, new[] { new OfficeRichTextParagraph(explicitStyle
            ? new[] { new OfficeRichTextRun("", 14, OfficeColor.Blue, italic: true, fontFamily: "Times New Roman") }
            : Array.Empty<OfficeRichTextRun>(), p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops) });
        PublisherDocument saved = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        PublisherTextStory story = saved.TextStories.Single();
        Assert.Equal(original.Id, story.Id); Assert.Equal("\n", story.Text); Assert.Single(story.Paragraphs);
        if (explicitStyle) {
            OfficeRichTextRun run = Assert.Single(story.Paragraphs[0].Runs);
            Assert.Equal(("Times New Roman", 14D, true, OfficeColor.Blue), (run.FontFamily, run.FontSize, run.Italic, run.Color));
        }
        Assert.Equal(document.Pages[0].TextFrames[0].Id, saved.Pages[0].TextFrames[0].Id);
    }

    private static IEnumerable<PublisherTable> Tables(PublisherDocument document) =>
        document.Pages.Concat(document.MasterPages).SelectMany(page => page.Tables);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Trailing_empty_paragraphs_use_the_retained_CR_or_the_requested_typing_style_during_count_growth(bool explicitStyle) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"));
        PublisherTextStory original = document.TextStories.Single(); OfficeRichTextParagraph p = original.Paragraphs[0];
        var empty = new OfficeRichTextParagraph(explicitStyle
            ? new[] { new OfficeRichTextRun("", 14, OfficeColor.Blue, italic: true, fontFamily: "Times New Roman") }
            : Array.Empty<OfficeRichTextRun>(), p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
        document.ReplaceStoryParagraphs(original.Id, new[] { p, empty });
        PublisherDocument saved = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        PublisherTextStory story = saved.TextStories.Single();
        Assert.Equal(original.Text + "\n", story.Text); Assert.Equal(2, story.Paragraphs.Count);
        Assert.Equal(p.Runs.Select(run => run.Text), story.Paragraphs[0].Runs.Select(run => run.Text));
        OfficeRichTextRun typing = Assert.Single(story.Paragraphs[1].Runs);
        Assert.Equal("", typing.Text);
        if (explicitStyle) Assert.Equal(("Times New Roman", 14D, true, OfficeColor.Blue),
            (typing.FontFamily, typing.FontSize, typing.Italic, typing.Color));
    }

    [Fact]
    public void Explicit_empty_typing_does_not_bypass_an_unsupported_CR_size_encoding() {
        const string fixture = "NativeComparison/TableCellTypingEarlyRange/native-tablecelltyping.pub";
        PublisherDocument original = PublisherDocument.Load(PublisherNativeTests.Fixture(fixture));
        PublisherTable table = Tables(original).First();
        PublisherQuillStoryStyles binding = original.SourceStyleCatalog.Stories.Single(story => story.StoryId == table.StoryId);
        int cr = binding.NativeStart + table.Cells[0].Text.Length;
        PublisherQuillSourceSpan span = binding.Uses.Single(use => use.Tag == "FDPC" && use.NativeStart <= cr
            && use.NativeStart + use.NativeLength > cr).DirectDefinition;
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            var data = new PublisherBinaryData(bytes, "Quill");
            var blocks = new PublisherBlockReader(data, new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None));
            PublisherBlock field = blocks.Chunk(span.Offset, span.Offset + span.Length).Single(item => item.Id == 0x12);
            bytes[field.Offset - 2] = 0x39;
            BitConverter.GetBytes(127000U).CopyTo(bytes, field.Offset);
        }, fixture);
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.ToPubBytes());
        OfficeRichTextParagraph p = Tables(document).First().Cells[0].Paragraphs[0];
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] { new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("", 14, OfficeColor.Blue, italic: true, fontFamily: "Times New Roman")
        }, p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops) });
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(
            new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains(error.Report.FidelityDiagnostics, item => item.LossKind == OfficeConversionLossKind.Failure);
    }
}
