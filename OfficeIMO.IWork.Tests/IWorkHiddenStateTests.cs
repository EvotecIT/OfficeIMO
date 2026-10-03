using OfficeIMO.IWork;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Hidden_state_selectors_require_partial_output_even_with_zero_model_counts(IWorkDocumentKind kind) {
        using MemoryStream package = HiddenStatePackage(kind, HiddenOwner(rowFields: BytesField(2, VarintField(2, 1))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        Assert.Equal(42d, Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells).Value);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false,
            new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.True(report.IsPartialEditableReconstruction);
        Assert.Throws<InvalidOperationException>(() => report.RequireCompleteEditableReconstruction());
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal("70/2[1]/3/2[1]/2", issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.UnsupportedField, issue.Kind);
        Assert.Equal(11ul, issue.Owner.RecordIdentifier);
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_HIDDEN_STATES_UNASSESSED"
            && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
    }

    [Theory]
    [InlineData(2, 2)]
    [InlineData(2, 3)]
    [InlineData(2, 4)]
    [InlineData(12, 2)]
    [InlineData(12, 3)]
    [InlineData(12, 4)]
    public void Base_and_summary_user_filtered_and_pivot_selectors_keep_physical_paths(int field, int flag) {
        using MemoryStream package = HiddenStatePackage(IWorkDocumentKind.Numbers,
            HiddenOwner(columnFields: BytesField(field, VarintField(flag, 1))));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.False(projection.HasEditableContent);
        Assert.Equal($"70/2[1]/2/{field}[1]/{flag}", Assert.Single(projection.SourceDeclarationIssues).FieldPath);
    }

    [Theory]
    [InlineData(5)]
    [InlineData(7)]
    [InlineData(9)]
    [InlineData(10)]
    [InlineData(11)]
    public void Unqualified_extent_payloads_keep_evidence_without_inferred_positions(int field) {
        using MemoryStream package = HiddenStatePackage(IWorkDocumentKind.Numbers,
            HiddenOwner(rowFields: BytesField(field, Message())));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.False(projection.HasEditableContent);
        Assert.Equal($"70/2[1]/3/{field}", Assert.Single(projection.SourceDeclarationIssues).FieldPath);
        Assert.Equal(42d, Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells).Value);
    }

    [Theory]
    [InlineData("duplicate", "70", IWorkSourceDeclarationIssueKind.RejectedMessageSet)]
    [InlineData("malformed", "70", IWorkSourceDeclarationIssueKind.MalformedMessage)]
    [InlineData("wrongStateWire", "70/2[1]", IWorkSourceDeclarationIssueKind.MalformedMessage)]
    [InlineData("missingAxis", "70/2[1]/3", IWorkSourceDeclarationIssueKind.MalformedMessage)]
    [InlineData("wrongDirection", "70/2[1]/3/3", IWorkSourceDeclarationIssueKind.InvalidValue)]
    [InlineData("ambiguousFlag", "70/2[1]/3/2[1]/2", IWorkSourceDeclarationIssueKind.InvalidValue)]
    public void Malformed_hidden_state_declarations_do_not_restore_complete_acceptance(string defect, string path,
        IWorkSourceDeclarationIssueKind expected) {
        byte[] owner = defect switch {
            "wrongStateWire" => VarintField(2, 1),
            "missingAxis" => BytesField(2, BytesField(2, VarintField(3, 0))),
            "wrongDirection" => HiddenOwner(rowDirection: 0),
            "ambiguousFlag" => HiddenOwner(rowFields: BytesField(2, Message(VarintField(2, 1), VarintField(2, 0)))),
            _ => HiddenOwner()
        };
        byte[] field = defect == "malformed" ? BytesField(70, new byte[] { 0x80 }) : BytesField(70, owner);
        if (defect == "duplicate") field = Message(field, field);
        using MemoryStream package = HiddenStatePackage(IWorkDocumentKind.Numbers, ownerField: field);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.False(projection.HasEditableContent);
        IWorkSourceDeclarationIssue issue = Assert.Single(projection.SourceDeclarationIssues);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal(expected, issue.Kind);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Both_model_and_extent_filter_routes_assess_enablement_and_skip_disabled_rules(bool modern) {
        foreach (bool enabled in new[] { false, true }) {
            using MemoryStream package = HiddenStatePackage(IWorkDocumentKind.Numbers,
                modern ? HiddenOwner(rowFields: ReferenceField(8, 30)) : HiddenOwner(),
                modelFields: modern ? Message() : ReferenceField(38, 30),
                records: ArchiveRecord(30, 6220, Message(VarintField(2, enabled ? 1ul : 0ul),
                    BytesField(7, new byte[] { 0x80 }), ReferenceField(3, 999))));
            IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
            Assert.Equal(!enabled, projection.HasEditableContent);
            Assert.Empty(projection.SourceReferenceIssues);
            if (!enabled) Assert.Empty(projection.SourceDeclarationIssues);
            else {
                IWorkSourceDeclarationIssue issue = Assert.Single(projection.SourceDeclarationIssues);
                Assert.Equal(30ul, issue.Owner.RecordIdentifier);
                Assert.Equal("2", issue.FieldPath);
                Assert.Equal(IWorkSourceDeclarationIssueKind.UnsupportedField, issue.Kind);
            }
        }
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("wrongType")]
    [InlineData("malformed")]
    [InlineData("missingFlag")]
    public void Unreadable_selected_filter_targets_remain_unassessed(string defect) {
        byte[] records = defect == "missing" ? Message()
            : ArchiveRecord(30, defect == "wrongType" ? 6001u : 6220u,
                defect == "malformed" ? new byte[] { 0x80 } : Message());
        using MemoryStream package = HiddenStatePackage(IWorkDocumentKind.Numbers,
            HiddenOwner(rowFields: ReferenceField(8, 30)), records: records);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.False(projection.HasEditableContent);
        if (defect == "missing") Assert.Equal("70/2[1]/3/8", Assert.Single(projection.SourceReferenceIssues).FieldPath);
        else Assert.Equal(30ul, Assert.Single(projection.SourceDeclarationIssues).Owner.RecordIdentifier);
    }

    [Fact]
    public void Empty_extents_and_false_selectors_keep_the_editable_contract() {
        using MemoryStream package = HiddenStatePackage(IWorkDocumentKind.Numbers,
            HiddenOwner(rowFields: Message(VarintField(6, 0), BytesField(2,
                Message(VarintField(2, 0), VarintField(3, 0), VarintField(4, 0))))));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.True(projection.HasEditableContent);
        Assert.Empty(projection.SourceDeclarationIssues);
    }

    [Fact]
    public void Hidden_state_entries_are_charged_before_nested_decode_and_parser_limits_remain_fatal() {
        using MemoryStream package = HiddenStatePackage(IWorkDocumentKind.Numbers,
            HiddenOwner(rowFields: BytesField(2, new byte[] { 0x80 })));
        Assert.Contains("dimension entries", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumTableDimensionEntries = 4 }).ReadNumbers()).Message);
        using MemoryStream depth = HiddenStatePackage(IWorkDocumentKind.Numbers,
            HiddenOwner(rowFields: BytesField(2, VarintField(2, 0))));
        Assert.Contains("depth", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(depth,
            new IWorkReadOptions { MaximumProtobufDepth = 3 }).ReadNumbers()).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Shared_hidden_states_deduplicate_evidence_and_distinct_selectors_use_the_existing_limit() {
        byte[] rowFields = BytesField(2, VarintField(2, 1));
        using MemoryStream shared = HiddenStatePackage(IWorkDocumentKind.Numbers, HiddenOwner(rowFields: rowFields), repeatModel: true);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(shared,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).ReadNumbers();
        Assert.Equal(2, Assert.Single(projection.Sheets).Tables.Count);
        Assert.Single(projection.SourceDeclarationIssues);
        using MemoryStream distinct = HiddenStatePackage(IWorkDocumentKind.Numbers,
            HiddenOwner(rowFields: Message(rowFields, rowFields)));
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(distinct,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).ReadNumbers()).Message);
    }

    [Fact]
    public void Hidden_state_partial_output_and_reader_keep_values_and_visibility_warning() {
        using MemoryStream package = HiddenStatePackage(IWorkDocumentKind.Numbers,
            HiddenOwner(rowFields: BytesField(2, VarintField(2, 1))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        using var automatic = source.ToExcelDocumentResult();
        Assert.True(automatic.IsVisualFallback);
        using var partial = source.ToExcelDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.True(partial.Report.IsPartialEditableReconstruction);
        using var saved = new MemoryStream(); partial.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        package.Position = 0;
        var result = new OfficeIMO.Reader.OfficeDocumentReaderBuilder().AddIWorkHandler().Build().ReadDocument(package, "hidden.numbers");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_HIDDEN_STATES_UNASSESSED");
        Assert.Equal("42", Assert.Single(Assert.Single(result.Tables).Rows[0]));
    }

    private static byte[] HiddenOwner(byte[]? rowFields = null, byte[]? columnFields = null, ulong rowDirection = 1) =>
        BytesField(2, Message(BytesField(2, Message(VarintField(3, 0), columnFields ?? Message())),
            BytesField(3, Message(VarintField(3, rowDirection), rowFields ?? Message()))));

    private static MemoryStream HiddenStatePackage(IWorkDocumentKind kind, byte[]? owner = null, byte[]? ownerField = null,
        byte[]? modelFields = null, byte[]? records = null, bool repeatModel = false) =>
        TableDependencyPackage(kind, Message(), additionalRecords: records, repeatModel: repeatModel, modelPayload: Message(
            BytesField(4, BytesField(3, BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 12))))),
            VarintField(6, 3), VarintField(7, 1), VarintField(14, 0), VarintField(15, 0),
            ownerField ?? BytesField(70, owner ?? HiddenOwner()), modelFields ?? Message()));
}
