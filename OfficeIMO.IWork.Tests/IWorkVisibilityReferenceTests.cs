using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> VisibilityDependencyKinds() {
        foreach (var kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            foreach (string route in new[] { "map", "model", "row", "column" }) yield return new object[] { kind, route };
    }

    [Theory]
    [MemberData(nameof(VisibilityDependencyKinds))]
    public void Selected_visibility_dependencies_keep_wrong_target_type_and_existing_declaration_evidence(
        IWorkDocumentKind kind, string route) {
        using var package = route == "map" ? MappedHiddenPackage(kind, mapType: 2021)
            : HiddenStatePackage(kind, HiddenOwner(
                rowFields: route == "row" ? ReferenceField(8, 30) : Message(),
                columnFields: route == "column" ? ReferenceField(8, 30) : Message()),
                modelFields: route == "model" ? ReferenceField(38, 30) : Message(),
                records: ArchiveRecord(30, 2021, VarintField(2, 0)));
        var (table, _) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        Assert.Equal(42d, Assert.Single(table.Cells).Value);
        Assert.Empty(table.HiddenRows);
        Assert.Empty(table.HiddenColumns);
        package.Position = 0;
        var report = ConvertUnitReport(package, kind, readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.SourceDeclarationIssues, item => item.Owner.RecordIdentifier == 30 && item.FieldPath == "$"
            && item.Kind == IWorkSourceDeclarationIssueKind.RejectedMessageSet);
        var issue = Assert.Single(report.SourceReferenceIssues);
        Assert.Equal(11ul, issue.Owner.RecordIdentifier);
        Assert.Equal(route switch { "map" => "46", "model" => "38", "column" => "70/2[1]/2/8", _ => "70/2[1]/3/8" }, issue.FieldPath);
        Assert.Equal(1, issue.ReferenceIndex);
        Assert.Equal(30ul, issue.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            item => item.Code == "IWORK_SOURCE_REFERENCES_UNRESOLVED").LossKind);
    }

    [Fact]
    public void Shared_wrong_type_position_map_retains_one_physical_reference_across_axes_and_models() {
        using var package = MappedHiddenPackage(IWorkDocumentKind.Numbers, mapType: 2021, repeatModel: true);
        var projection = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumSourceReferenceIssues = 1 }).ReadNumbers();
        Assert.Equal(2, Assert.Single(projection.Sheets).Tables.Count);
        Assert.Equal("46", Assert.Single(projection.SourceReferenceIssues).FieldPath);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    public void Distinct_filter_links_to_one_wrong_target_charge_physical_reference_budget(int limit) {
        using var package = HiddenStatePackage(IWorkDocumentKind.Numbers,
            HiddenOwner(rowFields: ReferenceField(8, 30), columnFields: ReferenceField(8, 30)),
            modelFields: ReferenceField(38, 30), records: ArchiveRecord(30, 2021, Message()), repeatModel: true);
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumSourceReferenceIssues = limit });
        if (limit == 2) {
            Assert.Contains("reference issues", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
            return;
        }
        var projection = source.ReadNumbers();
        Assert.Equal(2, Assert.Single(projection.Sheets).Tables.Count);
        Assert.Equal(new[] { "38", "70/2[1]/2/8", "70/2[1]/3/8" }, projection.SourceReferenceIssues.Select(item => item.FieldPath));
        Assert.All(projection.SourceReferenceIssues, item => {
            Assert.Equal(30ul, item.TargetIdentifier);
            Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, item.Kind);
        });
        Assert.Single(projection.SourceDeclarationIssues);
    }

    [Fact]
    public void Inactive_position_map_stays_outside_visibility_reference_inventory() {
        using var package = HiddenStatePackage(IWorkDocumentKind.Numbers,
            modelFields: ReferenceField(46, 30), records: ArchiveRecord(30, 2021, Message()));
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.True(projection.HasEditableContent);
        Assert.Empty(projection.SourceReferenceIssues);
        Assert.Empty(projection.SourceDeclarationIssues);
    }
}
