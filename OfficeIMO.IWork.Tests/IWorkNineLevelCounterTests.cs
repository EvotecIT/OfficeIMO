using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    private static readonly int[] NineLevelSourceLevels = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 8, 4, 5, 6, 7, 8, 0, 1, 2, 3, 4, 5, 6, 7, 8 };
    private static readonly int[] NineLevelSourceNumbers = { 7, 1, 1, 1, 1, 1, 1, 1, 5, 6, 2, 1, 1, 1, 1, 8, 1, 1, 1, 1, 1, 1, 1, 1 };

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Nine_level_source_counters_reset_through_the_full_shared_native_vector(IWorkDocumentKind kind) {
        using var package = NineLevelCounterPackage(kind);
        var content = ReadSelectedRichTable(IWorkSourceDocument.Open(package), kind).Item1.Cells.Single().RichText!;
        Assert.True(content.IsFormattingComplete);
        Assert.Equal(NineLevelSourceLevels, content.Paragraphs.Select(p => p.ListLevel));
        Assert.Equal(NineLevelSourceNumbers, content.Paragraphs.Select(p => p.ListNumber!.Value));
    }

    [Fact]
    public void Saved_PPTX_retains_all_nine_levels_and_deep_counters_through_the_existing_drawing_owner() {
        using var package = NineLevelCounterPackage(IWorkDocumentKind.Keynote);
        using var result = IWorkSourceDocument.Open(package).ToPowerPointDocumentResult(new() {
            Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        var content = Assert.Single(Assert.Single(result.Projection.Slides).Tables).Cells[0].RichText!;
        Assert.True(content.IsFormattingComplete);
        result.Value.SlideSize.SetSize(OfficeLength.Points(600), OfficeLength.Points(600));
        var table = Assert.Single(Assert.Single(result.Value.Slides).Tables);
        table.Width = OfficeLength.Points(500); table.Height = OfficeLength.Points(500);
        table.SetColumnWidth(0, OfficeIMO.OfficeLength.Points(500)); table.SetRowHeight(0, OfficeIMO.OfficeLength.Points(500));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        Assert.Empty(reopened.ValidateDocument());
        var paragraphs = Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0).Paragraphs;
        Assert.Equal(NineLevelSourceLevels.Select(level => (int?)level), paragraphs.Select(p => p.Level));
        string[] expected = content.Paragraphs.Select(p => p.ListNumber + ". " + p.Text).ToArray();
        Assert.Equal(string.Concat(expected), DrawingText(reopened.Slides[0].CreateVisualSnapshot().Drawing));
    }

    [Theory]
    [InlineData("missing-slot")]
    [InlineData("tiered")]
    [InlineData("nondecimal")]
    [InlineData("mixed-marker")]
    public void The_deepest_decimal_profile_requires_every_selected_ancestor_slot(string fault) {
        byte[] style = Message(Enumerable.Range(0, fault == "missing-slot" ? 8 : 9)
            .Select(level => Message(VarintField(11, fault == "mixed-marker" && level == 7 ? 2ul : 3ul),
                VarintField(15, fault == "nondecimal" && level == 7 ? 1ul : 0ul),
                VarintField(25, fault == "tiered" && level == 7 ? 1ul : 0ul))).ToArray());
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            Message(AttributeTable(6, ListLevelEntry(0, 8)), AttributeTable(14, ListLevelEntry(0, 7)),
                AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19)))),
            additionalRecords: ArchiveRecord(19, 2023, style));
        var content = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells.Single().RichText!;
        Assert.False(content.IsFormattingComplete); Assert.Null(content.Paragraphs.Single().ListNumber);
        Assert.Equal("Value", content.Paragraphs.Single().Text);
    }

    private static MemoryStream NineLevelCounterPackage(IWorkDocumentKind kind) {
        string[] text = Enumerable.Range(0, NineLevelSourceLevels.Length).Select(index => index == 0 ? "Value" : "Item " + index).ToArray();
        var offsets = ParagraphOffsets(text);
        return SelectedRichTextPackage(kind,
            Message(StringField(3, "\n" + string.Join("\n", text.Skip(1))),
                AttributeTable(6, offsets.Select((offset, index) => ListLevelEntry(offset,
                    (ulong)NineLevelSourceLevels[index], index is 0 or 8 ? 1ul : 0ul)).ToArray()),
                AttributeTable(14, offsets.Select((offset, index) => ListLevelEntry(offset, index == 0 ? 7ul : index == 8 ? 5ul : 0ul)).ToArray()),
                AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19))),
                AttributeTable(5, AttributeEntry(0, ReferenceField(2, 31))),
                AttributeTable(8, AttributeEntry(0, ReferenceField(2, 33)))),
            additionalRecords: Message(
                ArchiveRecord(19, 2023, Message(Enumerable.Range(0, 9).Select(level => Message(
                    VarintField(11, 3), VarintField(15, 0), VarintField(25, 0), FloatField(12, 1.5f), FloatField(13, level * 18),
                    BytesField(14, Message(FloatField(1, 1), FloatField(2, 0), VarintField(3, 1))))).ToArray())),
                ArchiveRecord(31, 2022, BytesField(11, Message(FloatField(3, 12), StringField(5, "Arial")))),
                ArchiveRecord(33, 2021, BytesField(11, Message(FloatField(3, 12), StringField(5, "Arial"))))));
    }
}
