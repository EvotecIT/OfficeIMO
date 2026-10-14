using OfficeIMO.IWork;
using OfficeIMO.Reader.IWork;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Decimal_child_counters_reset_after_parents_and_continue_across_plain_text(IWorkDocumentKind kind) {
        string[] text = { "Value", "Child one", "Child two", "Parent eight", "Child reset", "Gap",
            "Child continued", "Child five", "Parent three", "Parent four" };
        ulong[] offsets = ParagraphOffsets(text);
        using var package = SelectedRichTextPackage(kind,
            Message(StringField(3, "\n" + string.Join("\n", text.Skip(1))),
                AttributeTable(6, ListLevelEntry(0, 0, 1), ListLevelEntry(offsets[1], 1),
                    ListLevelEntry(offsets[3], 0), ListLevelEntry(offsets[4], 1),
                    ListLevelEntry(offsets[5], 0), ListLevelEntry(offsets[6], 1),
                    ListLevelEntry(offsets[7], 1, 1), ListLevelEntry(offsets[8], 0, 1),
                    ListLevelEntry(offsets[9], 0)),
                AttributeTable(14, ListLevelEntry(0, 7), ListLevelEntry(offsets[1], 0),
                    ListLevelEntry(offsets[7], 5), ListLevelEntry(offsets[8], 3),
                    ListLevelEntry(offsets[9], 0)),
                AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19)),
                    AttributeEntry(offsets[5], ReferenceField(2, 22)),
                    AttributeEntry(offsets[6], ReferenceField(2, 19)))),
            additionalRecords: Message(
                ArchiveRecord(19, 2023, Message(VarintField(11, 3), VarintField(11, 3),
                    VarintField(15, 0), VarintField(15, 0), VarintField(25, 0), VarintField(25, 0))),
                ArchiveRecord(22, 2023, VarintField(11, 0))));
        var content = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), kind).Item1.Cells).RichText!;
        Assert.True(content.IsFormattingComplete);
        Assert.Equal(new int?[] { 7, 1, 2, 8, 1, null, 2, 5, 3, 4 },
            content.Paragraphs.Select(paragraph => paragraph.ListNumber));
        Assert.Equal(new int?[] { 7, null, null, null, null, null, null, 5, 3, null },
            content.Paragraphs.Select(paragraph => paragraph.ListStartNumber));
        Assert.Equal("  2. Child continued", IWorkReadProjection.RichTextMarkdown(content.Paragraphs[6]));
        if (kind == IWorkDocumentKind.Keynote) {
            package.Position = 0;
            using var result = IWorkSourceDocument.Open(package).ToPowerPointPresentationResult(new() {
                Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
            using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
            using var reopened = PowerPointPresentation.Load(saved);
            var paragraphs = Assert.Single(Assert.Single(reopened.Slides).Tables).GetCell(0, 0).Paragraphs;
            Assert.Equal(new int?[] { 7, 1, null, 8, null, null, 2, 5, 3, null },
                paragraphs.Select(paragraph => paragraph.NumberingStartAt));
            Assert.Equal(new int?[] { 0, 1, 1, 0, 1, null, 1, 1, 0, 0 }, paragraphs.Select(paragraph => paragraph.Level));
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    [Theory]
    [InlineData("unknown-flag")]
    [InlineData("restart-with-zero")]
    [InlineData("restart-on-non-list")]
    [InlineData("restart-inside-paragraph")]
    [InlineData("restart-at-end")]
    public void Rejected_restart_flags_never_silently_qualify_continuation(string fault) {
        byte[] levels = fault switch {
            "unknown-flag" => AttributeTable(6, ListLevelEntry(0, 0, 2)),
            "restart-inside-paragraph" => AttributeTable(6, ListLevelEntry(1, 0, 1)),
            "restart-at-end" => AttributeTable(6, ListLevelEntry(27, 0, 1)),
            _ => AttributeTable(6, ListLevelEntry(0, 0, 1))
        };
        byte[] starts = fault is "restart-with-zero" or "restart-at-end"
            ? AttributeTable(14, ListLevelEntry(0, 0))
            : AttributeTable(14, ListLevelEntry(0, 7), ListLevelEntry(6, 0));
        byte[]? styles = fault == "restart-on-non-list"
            ? AttributeTable(7, AttributeEntry(0, ReferenceField(2, 22))) : null;
        using var package = DecimalCounterPackage(IWorkDocumentKind.Numbers, starts,
            listStyles: styles, levels: levels);
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        var content = Assert.Single(Assert.Single(projection.Sheets).Tables).Cells[0].RichText!;
        Assert.False(content.IsFormattingComplete);
        Assert.Contains(projection.SourceDeclarationIssues, issue => issue.FieldPath.StartsWith("6/1[", StringComparison.Ordinal));
        Assert.Null(content.Paragraphs[0].ListNumber);
        Assert.All(content.Paragraphs.Skip(1), paragraph => Assert.Null(paragraph.ListNumber));
    }

    private static ulong[] ParagraphOffsets(IReadOnlyList<string> text) {
        var offsets = new ulong[text.Count];
        for (int index = 1; index < text.Count; index++) offsets[index] = offsets[index - 1] + (ulong)text[index - 1].Length + 1;
        return offsets;
    }

    [Fact]
    public void A_rejected_plain_list_style_cannot_bridge_a_counter_gap() {
        using var package = DecimalCounterPackage(IWorkDocumentKind.Numbers,
            starts: AttributeTable(14, ListLevelEntry(0, 7), ListLevelEntry(6, 0)),
            listStyles: AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19)),
                AttributeEntry(6, ReferenceField(2, 22)), AttributeEntry(11, ReferenceField(2, 19))),
            plainStyle: Message(VarintField(11, 0), FloatField(11, 1)));
        var content = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package),
            IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.False(content.IsFormattingComplete);
        Assert.Equal(new int?[] { 7, null, null, null }, content.Paragraphs.Select(paragraph => paragraph.ListNumber));
    }
}
