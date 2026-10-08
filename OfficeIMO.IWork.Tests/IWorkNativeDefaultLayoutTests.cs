using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Native_empty_layout_declarations_reset_inherited_spacing_and_custom_tabs(IWorkDocumentKind kind) {
        using var package = SelectedRichTextPackage(kind,
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))), additionalRecords: Message(
                ArchiveRecord(19, 2022, Message(BytesField(1, ReferenceField(3, 20)),
                    BytesField(12, Message(BytesField(13, Message()), BytesField(25, Message()))))),
                ArchiveRecord(20, 2022, BytesField(12, Message(BytesField(13, FloatField(2, 2)),
                    BytesField(25, BytesField(1, FloatField(1, 18))))))));
        var text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells).RichText!;
        Assert.True(text.IsFormattingComplete);
        var paragraph = Assert.Single(text.Paragraphs);
        Assert.Equal(1, paragraph.Style.LineSpacingMultiplier);
        Assert.Empty(paragraph.Style.TabStops!);
        package.Position = 0;
        Assert.Empty(ConvertUnitReport(package, kind, visual: false).SourceDeclarationIssues);
    }

    [Fact]
    public void Native_empty_layout_reset_does_not_hide_unassessed_ancestor_declarations() {
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))), additionalRecords: Message(
                ArchiveRecord(19, 2022, Message(BytesField(1, ReferenceField(3, 20)),
                    BytesField(12, Message(BytesField(13, Message()), BytesField(25, Message()))))),
                ArchiveRecord(20, 2022, BytesField(12, BytesField(13, Message(VarintField(1, 2), FloatField(2, 12)))))));
        var source = IWorkSourceDocument.Open(package);
        Assert.False(Assert.Single(ReadSelectedRichTable(source, IWorkDocumentKind.Numbers).Item1.Cells).RichText!.IsFormattingComplete);
        package.Position = 0;
        var issue = Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false).SourceDeclarationIssues);
        Assert.Equal(20ul, issue.Owner.RecordIdentifier);
        Assert.Equal("12/13", issue.FieldPath);
    }
}
