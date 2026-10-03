using OfficeIMO.IWork;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Explicit_list_levels_override_equal_indents_and_keep_cache_entries_distinct(IWorkDocumentKind kind) {
        using MemoryStream package = ListLevelPackage(kind, AttributeTable(6,
            ListLevelEntry(0, 1), ListLevelEntry(6, 0)));
        var content = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), kind).Item1.Cells).RichText!;
        Assert.Equal(new[] { 1, 0 }, content.Paragraphs.Select(p => p.ListLevel));
        Assert.Equal(new[] { "–", "•" }, content.Paragraphs.Select(p => p.ListLabel));
        Assert.True(content.IsFormattingComplete);
        if (kind != IWorkDocumentKind.Pages) return;
        package.Position = 0;
        using var result = IWorkSourceDocument.Open(package).ToWordDocumentResult(
            new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true });
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordprocessingDocument.Open(saved, false);
        var paragraphs = reopened.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>()
            .Where(p => p.InnerText == "Value" || p.InnerText == "Next").ToArray();
        Assert.Equal(new[] { 1, 0 }, paragraphs.Select(p => p.ParagraphProperties!.NumberingProperties!
            .NumberingLevelReference!.Val!.Value));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(reopened));
    }

    [Fact]
    public void Invalid_level_boundary_does_not_carry_the_previous_explicit_level() {
        using MemoryStream package = ListLevelPackage(IWorkDocumentKind.Numbers, AttributeTable(6,
            ListLevelEntry(0, 1), Message(VarintField(1, 6), FloatField(2, 1), VarintField(3, 0))));
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        var content = Assert.Single(Assert.Single(projection.Sheets).Tables).Cells[0].RichText!;
        Assert.False(content.IsFormattingComplete);
        Assert.Equal(new[] { 1, 0 }, content.Paragraphs.Select(p => p.ListLevel));
        Assert.Contains(projection.SourceDeclarationIssues, issue => issue.Owner.RecordIdentifier == 15
            && issue.FieldPath == "6/1[2]/2" && issue.Kind == IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
    }

    [Fact]
    public void Unqualified_paragraph_counter_metadata_remains_reported() {
        using MemoryStream package = ListLevelPackage(IWorkDocumentKind.Numbers,
            AttributeTable(6, ListLevelEntry(0, 1, 1)));
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.Contains(projection.SourceDeclarationIssues, issue => issue.Owner.RecordIdentifier == 15
            && issue.FieldPath == "6/1[1]/3" && issue.Kind == IWorkSourceDeclarationIssueKind.UnsupportedField);
        Assert.False(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells[0].RichText!.IsFormattingComplete);
    }

    [Fact]
    public void Explicit_list_level_boundaries_charge_the_text_attribute_budget() {
        using MemoryStream package = ListLevelPackage(IWorkDocumentKind.Numbers,
            AttributeTable(6, Enumerable.Range(0, 20).Select(i => ListLevelEntry(0, 0)).ToArray()));
        var source = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumProjectedTextItems = 10 });
        Assert.Contains("attribute", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message,
            StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] ListLevelEntry(ulong offset, ulong level, ulong value = 0) =>
        Message(VarintField(1, offset), VarintField(2, level), VarintField(3, value));

    private static MemoryStream ListLevelPackage(IWorkDocumentKind kind, byte[] attributes) =>
        SelectedRichTextPackage(kind, Message(StringField(3, "\nNext"), attributes,
            AttributeTable(7, AttributeEntry(0, ReferenceField(2, 19))),
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 21)))),
            additionalRecords: Message(
                ArchiveRecord(19, 2023, Message(VarintField(11, 2), VarintField(11, 2),
                    FloatField(13, 0), FloatField(13, 0), StringField(16, "•"), StringField(16, "–"))),
                ArchiveRecord(21, 2022, BytesField(12, FloatField(11, 0)))));
}
