using OfficeIMO.IWork;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Declared_hidden_content_requires_partial_policy_in_each_destination(IWorkDocumentKind kind) {
        using MemoryStream package = VisibilityCountPackage(kind, VarintField(14, 1));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Equal(42d, Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells).Value);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false,
            new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.True(report.IsPartialEditableReconstruction);
        Assert.Throws<InvalidOperationException>(() => report.RequireCompleteEditableReconstruction());
        Assert.Empty(report.PreservedRecords);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(11ul, issue.Owner.RecordIdentifier);
        Assert.Equal("14", issue.FieldPath);
        Assert.Equal(1, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.UnsupportedField, issue.Kind);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_VISIBILITY_UNASSESSED"
            && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
    }

    [Theory]
    [InlineData(14)]
    [InlineData(15)]
    [InlineData(40)]
    [InlineData(41)]
    [InlineData(42)]
    public void Each_native_hidden_or_filtered_count_retains_its_own_path(int field) {
        using MemoryStream package = VisibilityCountPackage(IWorkDocumentKind.Numbers, VarintField(field, 1));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.False(projection.HasEditableContent);
        IWorkSourceDeclarationIssue issue = Assert.Single(projection.SourceDeclarationIssues);
        Assert.Equal(field.ToString(System.Globalization.CultureInfo.InvariantCulture), issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.UnsupportedField, issue.Kind);
    }

    [Theory]
    [InlineData("wrongWire", 1)]
    [InlineData("duplicate", 2)]
    [InlineData("huge", 1)]
    [InlineData("outOfBounds", 1)]
    public void Ambiguous_or_invalid_visibility_counts_do_not_claim_all_content_is_visible(string defect, int count) {
        byte[] declaration = defect switch {
            "wrongWire" => BytesField(14, new byte[] { 0 }),
            "duplicate" => Message(VarintField(14, 0), VarintField(14, 0)),
            "huge" => VarintField(14, ulong.MaxValue),
            _ => VarintField(14, 4)
        };
        using MemoryStream package = VisibilityCountPackage(IWorkDocumentKind.Numbers, declaration);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.False(projection.HasEditableContent);
        IWorkSourceDeclarationIssue issue = Assert.Single(projection.SourceDeclarationIssues);
        Assert.Equal("14", issue.FieldPath);
        Assert.Equal(count, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidValue, issue.Kind);
    }

    [Fact]
    public void Absent_or_explicit_zero_counts_keep_the_existing_editable_contract() {
        foreach (byte[] counts in new[] { Message(), Message(VarintField(14, 0), VarintField(15, 0),
            VarintField(40, 0), VarintField(41, 0), VarintField(42, 0)) }) {
            using MemoryStream package = VisibilityCountPackage(IWorkDocumentKind.Numbers, counts);
            IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
            Assert.True(projection.HasEditableContent);
            Assert.Empty(projection.SourceDeclarationIssues);
        }
    }

    [Fact]
    public void Visibility_evidence_deduplicates_shared_models_and_uses_the_captured_declaration_budget() {
        using MemoryStream package = VisibilityCountPackage(IWorkDocumentKind.Numbers, VarintField(14, 1), repeatModel: true);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).ReadNumbers();
        Assert.Equal(2, Assert.Single(projection.Sheets).Tables.Count);
        Assert.Single(projection.SourceDeclarationIssues);
        using MemoryStream overBudget = VisibilityCountPackage(IWorkDocumentKind.Numbers,
            Message(VarintField(14, 1), VarintField(41, 1)));
        var options = new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 };
        IWorkSourceDocument source = IWorkSourceDocument.Open(overBudget, options);
        options.MaximumSourceDeclarationIssues = 10;
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Partial_visibility_output_preserves_values_and_reader_warns_that_visibility_is_unassessed() {
        using MemoryStream package = VisibilityCountPackage(IWorkDocumentKind.Numbers, VarintField(14, 1));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        using var automatic = source.ToExcelDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(automatic.IsVisualFallback);
        using var partial = source.ToExcelDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.True(partial.Report.IsPartialEditableReconstruction);
        using var saved = new MemoryStream(); partial.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        package.Position = 0;
        var reader = new OfficeIMO.Reader.OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        var result = reader.ReadDocument(package, "visibility.numbers");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_VISIBILITY_UNASSESSED");
        Assert.Equal("42", Assert.Single(Assert.Single(result.Tables).Rows[0]));
    }

    private static MemoryStream VisibilityCountPackage(IWorkDocumentKind kind, byte[] counts, bool repeatModel = false) =>
        TableDependencyPackage(kind, Message(), repeatModel: repeatModel, modelPayload: Message(
            BytesField(4, BytesField(3, BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 12))))),
            VarintField(6, 3), VarintField(7, 1), counts));
}
