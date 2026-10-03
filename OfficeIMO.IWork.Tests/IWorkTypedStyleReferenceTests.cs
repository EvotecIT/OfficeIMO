using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> WrongTypeStyleReferences() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote }) {
            foreach (int field in new[] { 5, 7, 8, 11 }) yield return new object[] { kind, field, false };
            foreach (int field in new[] { 5, 7, 8 }) yield return new object[] { kind, field, true };
        }
    }

    [Theory]
    [MemberData(nameof(WrongTypeStyleReferences))]
    public void Selected_wrong_type_text_dependencies_keep_physical_reference_evidence_and_visible_text(
        IWorkDocumentKind kind, int field, bool parent) {
        uint expectedType = field switch { 5 => 2022u, 7 => 2023u, 8 => 2021u, _ => 2032u };
        uint wrongType = expectedType == 2021 ? 2023u : 2021u;
        byte[] records = parent
            ? Message(ArchiveRecord(19, expectedType, BytesField(1, ReferenceField(3, 20))),
                ArchiveRecord(20, wrongType, Message()))
            : ArchiveRecord(19, wrongType, Message());
        using var package = SelectedRichTextPackage(kind, AttributeTable(field, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: records);
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        IWorkTextContent text = Assert.Single(table.Cells).RichText!;
        Assert.Equal("Value", text.PlainText);
        Assert.False(text.IsFormattingComplete);
        IWorkSourceReferenceIssue issue = Assert.Single(issues);
        Assert.Equal(parent ? 19ul : 15ul, issue.Owner.RecordIdentifier);
        Assert.Equal(parent ? "1/3" : field + "/1[1]/2", issue.FieldPath);
        Assert.Equal(parent ? 20ul : 19ul, issue.TargetIdentifier);
        Assert.Equal(1, issue.ReferenceIndex);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        Assert.Single(report.SourceReferenceIssues);
        Assert.Empty(report.SourceDeclarationIssues);
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            diagnostic => diagnostic.Code == "IWORK_SOURCE_REFERENCES_UNRESOLVED").LossKind);
    }

    [Fact]
    public void Shared_wrong_type_text_references_deduplicate_and_remain_subject_to_the_reference_budget() {
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))), aliases: true,
            additionalRecords: ArchiveRecord(19, 2023, Message()));
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 1 }), IWorkDocumentKind.Numbers);
        Assert.Equal(2, table.Cells.Count);
        Assert.Single(issues);
        Assert.All(table.Cells, cell => Assert.False(cell.RichText!.IsFormattingComplete));

        using var repeated = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19)), AttributeEntry(5, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2023, Message()));
        IWorkSourceDocument source = IWorkSourceDocument.Open(repeated,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 1 });
        Assert.Contains("source reference issues", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Character_attributes_can_select_paragraph_styles_and_inherit_character_properties(IWorkDocumentKind kind) {
        using var package = SelectedRichTextPackage(kind, AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2022, BytesField(11, VarintField(1, 1))));
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        IWorkTextContent text = Assert.Single(table.Cells).RichText!;
        Assert.True(text.IsFormattingComplete);
        Assert.True(Assert.Single(Assert.Single(text.Paragraphs).Runs).Style.Bold);
        Assert.Empty(issues);
    }
}
