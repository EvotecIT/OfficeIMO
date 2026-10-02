using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> CatalogKinds() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            foreach (int field in new[] { 4, 6, 17 }) yield return new object[] { kind, field };
    }

    [Theory]
    [MemberData(nameof(CatalogKinds))]
    public void Malformed_selected_catalog_root_retains_owner_and_unknown_outer_count(IWorkDocumentKind kind, int field) {
        using MemoryStream package = field == 17
            ? SelectedRichTextPackage(kind, catalogPayload: new byte[] { 0x80 })
            : SelectedCatalogPackage(kind, field, new byte[] { 0x80 });
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(13ul, issue.Owner.RecordIdentifier);
        Assert.Equal("$", issue.FieldPath);
        Assert.Null(issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.MalformedMessage, issue.Kind);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 4)]
    [InlineData(IWorkDocumentKind.Numbers, 4)]
    [InlineData(IWorkDocumentKind.Keynote, 4)]
    [InlineData(IWorkDocumentKind.Pages, 6)]
    [InlineData(IWorkDocumentKind.Numbers, 6)]
    [InlineData(IWorkDocumentKind.Keynote, 6)]
    public void Malformed_duplicate_catalog_value_cannot_leave_a_trusted_first_value(IWorkDocumentKind kind, int field) {
        uint key = field == 4 ? 1u : 0u;
        byte[] first = Message(VarintField(1, key), field == 4 ? StringField(3, "First") : BytesField(5, FormulaConstant(1d)));
        byte[] second = Message(VarintField(1, key), field == 4 ? BytesField(3, new byte[] { 0xFF }) : BytesField(5, new byte[] { 0x80 }));
        using MemoryStream package = SelectedCatalogPackage(kind, field, Message(BytesField(3, first), BytesField(3, second)));
        (IWorkTable table, _) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        IWorkTableCell cell = Assert.Single(table.Cells);
        if (field == 4) {
            Assert.Equal(IWorkCellKind.Error, cell.Kind);
            Assert.Null(cell.Value);
        } else {
            Assert.Equal(IWorkCellKind.Formula, cell.Kind);
            Assert.False(cell.FormulaIsComplete);
            Assert.Equal(42d, cell.Value);
        }
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 4)]
    [InlineData(IWorkDocumentKind.Numbers, 4)]
    [InlineData(IWorkDocumentKind.Keynote, 4)]
    [InlineData(IWorkDocumentKind.Pages, 6)]
    [InlineData(IWorkDocumentKind.Numbers, 6)]
    [InlineData(IWorkDocumentKind.Keynote, 6)]
    public void Unreadable_value_with_distinct_key_keeps_healthy_catalog_value(IWorkDocumentKind kind, int field) {
        byte[] healthy = CatalogTestEntry(field);
        byte[] broken = Message(VarintField(1, 99), BytesField(field == 4 ? 3 : 5, new byte[] { 0xFF }));
        using MemoryStream package = SelectedCatalogPackage(kind, field,
            Message(BytesField(3, broken), BytesField(3, healthy)));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells);
        if (field == 4) Assert.Equal("First", cell.Value);
        else {
            Assert.True(cell.FormulaIsComplete);
            Assert.Equal(42d, cell.Value);
        }
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(13ul, issue.Owner.RecordIdentifier);
        Assert.Equal(field == 4 ? "3[1]/3" : "3[1]/5", issue.FieldPath);
        Assert.Equal(1, issue.DeclaredValueCount);
        Assert.Equal(field == 4 ? IWorkSourceDeclarationIssueKind.InvalidValue
            : IWorkSourceDeclarationIssueKind.MalformedMessage, issue.Kind);
        Assert.True(report.IsPartialEditableReconstruction);
    }

    [Theory]
    [MemberData(nameof(CatalogKinds))]
    public void Unreadable_catalog_key_does_not_establish_uniqueness(IWorkDocumentKind kind, int field) {
        using MemoryStream package = CatalogTestPackage(kind, field,
            Message(BytesField(3, Message()), BytesField(3, CatalogTestEntry(field))));
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells);
        if (field == 6) {
            Assert.False(cell.FormulaIsComplete);
            Assert.Equal(42d, cell.Value);
        } else {
            Assert.Equal(IWorkCellKind.Error, cell.Kind);
            Assert.Null(cell.Value);
        }
        package.Position = 0;
        IWorkSourceDeclarationIssue issue = Assert.Single(ConvertUnitReport(package, kind).SourceDeclarationIssues);
        Assert.Equal("3[1]/1", issue.FieldPath);
        Assert.Equal(0, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata, issue.Kind);
    }

    [Theory]
    [MemberData(nameof(CatalogKinds))]
    public void Third_duplicate_cannot_restore_catalog_key_or_hide_physical_evidence(IWorkDocumentKind kind, int field) {
        byte[] entry = BytesField(3, CatalogTestEntry(field));
        using MemoryStream package = CatalogTestPackage(kind, field, Message(entry, entry, entry));
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells);
        if (field == 6) Assert.False(cell.FormulaIsComplete);
        else Assert.Equal(IWorkCellKind.Error, cell.Kind);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        Assert.Equal(new[] { "3[2]/1", "3[3]/1" }, report.SourceDeclarationIssues.Select(issue => issue.FieldPath));
        Assert.All(report.SourceDeclarationIssues, issue => {
            Assert.Equal(13ul, issue.Owner.RecordIdentifier);
            Assert.Equal(1, issue.DeclaredValueCount);
            Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata, issue.Kind);
        });
    }

    [Theory]
    [MemberData(nameof(CatalogKinds))]
    public void Unsupported_catalog_envelope_retains_unknown_outer_count(IWorkDocumentKind kind, int field) {
        using MemoryStream package = CatalogTestPackage(kind, field,
            Message(VarintField(99, 0), BytesField(3, CatalogTestEntry(field))));
        IWorkSourceDeclarationIssue issue = Assert.Single(ConvertUnitReport(package, kind).SourceDeclarationIssues);
        Assert.Equal("$", issue.FieldPath);
        Assert.Null(issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.RejectedMessageSet, issue.Kind);
    }

    [Theory]
    [MemberData(nameof(CatalogKinds))]
    public void Configured_nested_catalog_field_limit_remains_fatal(IWorkDocumentKind kind, int field) {
        using MemoryStream package = CatalogTestPackage(kind, field,
            BytesField(3, Message(Enumerable.Range(0, 9).Select(_ => VarintField(1, 0)).ToArray())));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumProtobufFieldCount = 8 });
        Assert.Contains("field limit", Assert.Throws<InvalidDataException>(() => ReadSelectedRichTable(source, kind)).Message);
    }

    [Theory]
    [InlineData(4)]
    [InlineData(6)]
    [InlineData(17)]
    public void Unreadable_catalog_entries_consume_catalog_budget(int field) {
        using MemoryStream package = CatalogTestPackage(IWorkDocumentKind.Numbers, field,
            Message(BytesField(3, new byte[] { 0x80 }), BytesField(3, new byte[] { 0x80 })));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumTableCatalogEntries = 1 });
        Assert.Contains("catalog limit", Assert.Throws<InvalidDataException>(() => ReadSelectedRichTable(source, IWorkDocumentKind.Numbers)).Message);
    }

    [Theory]
    [InlineData(4)]
    [InlineData(6)]
    [InlineData(17)]
    public void Catalog_evidence_budget_is_cumulative_and_snapshotted(int field) {
        using MemoryStream package = CatalogTestPackage(IWorkDocumentKind.Numbers, field,
            Message(BytesField(3, new byte[] { 0x80 }), BytesField(3, new byte[] { 0x80 })));
        var options = new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 };
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers, options);
        options.MaximumSourceDeclarationIssues = 10;
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() => ReadSelectedRichTable(source, IWorkDocumentKind.Numbers)).Message);
    }

    [Fact]
    public void Reused_catalog_keeps_each_physical_evidence_path_once() {
        using MemoryStream package = SelectedCatalogPackage(IWorkDocumentKind.Numbers, 4,
            BytesField(3, Message(VarintField(1, 1), BytesField(3, new byte[] { 0xFF }))), repeatModel: true);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).ReadNumbers();
        Assert.Equal(2, Assert.Single(projection.Sheets).Tables.Count);
        Assert.Single(projection.SourceDeclarationIssues);
    }

    [Theory]
    [MemberData(nameof(CatalogKinds))]
    public void Unreadable_catalog_entries_retain_physical_positions_and_block_ambiguous_values(IWorkDocumentKind kind, int field) {
        using MemoryStream package = CatalogTestPackage(kind, field,
            Message(BytesField(3, new byte[] { 0x80 }), BytesField(3, CatalogTestEntry(field)), VarintField(3, 0)));
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells);
        if (field == 6) {
            Assert.False(cell.FormulaIsComplete);
            Assert.Equal(42d, cell.Value);
        } else Assert.Equal(IWorkCellKind.Error, cell.Kind);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Equal(new[] { "3[1]", "3[3]" }, report.SourceDeclarationIssues.Select(issue => issue.FieldPath));
        Assert.All(report.SourceDeclarationIssues, issue => {
            Assert.Equal(13ul, issue.Owner.RecordIdentifier);
            Assert.Equal(1, issue.DeclaredValueCount);
        });
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.FidelityDiagnostics, issue => issue.Code == "IWORK_SOURCE_DECLARATIONS_UNASSESSED");
    }

    [Fact]
    public void Ambiguous_string_values_still_consume_decoding_text_budget() {
        byte[] entry = BytesField(3, Message(VarintField(1, 1), StringField(3, new string('X', 40))));
        using MemoryStream package = SelectedCatalogPackage(IWorkDocumentKind.Numbers, 4, Message(entry, entry));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 50 });
        Assert.Contains("Text character count", Assert.Throws<InvalidDataException>(() => ReadSelectedRichTable(source, IWorkDocumentKind.Numbers)).Message);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 4)]
    [InlineData(IWorkDocumentKind.Numbers, 4)]
    [InlineData(IWorkDocumentKind.Keynote, 4)]
    [InlineData(IWorkDocumentKind.Pages, 6)]
    [InlineData(IWorkDocumentKind.Numbers, 6)]
    [InlineData(IWorkDocumentKind.Keynote, 6)]
    public void Partial_catalog_recovery_survives_destination_save_and_reopen(IWorkDocumentKind kind, int field) {
        byte[] broken = Message(VarintField(1, field == 4 ? 99u : 0u),
            BytesField(field == 4 ? 3 : 5, new byte[] { 0xFF }));
        using MemoryStream package = SelectedCatalogPackage(kind, field,
            Message(BytesField(3, CatalogTestEntry(field)), BytesField(3, broken)));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        string expected = field == 4 ? "First" : "42";
        if (kind == IWorkDocumentKind.Pages) {
            using var automatic = source.ToWordDocumentResult();
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToWordDocumentResult(options);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using WordDocument reopened = WordDocument.Load(saved);
            Assert.Equal(expected, reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var automatic = source.ToExcelDocumentResult();
            Assert.Equal(field == 4, automatic.IsVisualFallback);
            using var partial = source.ToExcelDocumentResult(options);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using ExcelDocument reopened = ExcelDocument.Load(saved);
            if (field == 4) Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
            else Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        } else {
            using var automatic = source.ToPowerPointPresentationResult();
            Assert.Equal(field == 4, automatic.IsVisualFallback);
            using var partial = source.ToPowerPointPresentationResult(options);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
            Assert.Equal(expected, reopened.Slides[0].Tables.First().GetCell(0, 0).Text);
        }
    }

    private static byte[] CatalogTestEntry(int field) => Message(VarintField(1, field == 6 ? 0u : 1u),
        field == 4 ? StringField(3, "First") : field == 6 ? BytesField(5, FormulaConstant(1d)) : ReferenceField(9, 14));

    private static MemoryStream CatalogTestPackage(IWorkDocumentKind kind, int field, byte[] payload) => field == 17
        ? SelectedRichTextPackage(kind, catalogPayload: payload) : SelectedCatalogPackage(kind, field, payload);

    private static MemoryStream SelectedCatalogPackage(IWorkDocumentKind kind, int field, byte[] payload,
        bool repeatModel = false, uint catalogType = 6005, bool includeCatalogKind = true) {
        byte[] cell = new byte[field == 4 ? 16 : 24]; cell[0] = 5; cell[1] = field == 4 ? (byte)3 : (byte)2;
        WriteUInt32(cell, 8, field == 4 ? 1u << 3 : (1u << 1) | (1u << 9));
        if (field == 4) WriteUInt32(cell, 12, 1);
        else Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, cell, 12, 8);
        return TableDependencyPackage(kind, ReferenceField(field, 13), cellPayload: cell,
            repeatModel: repeatModel, additionalRecords: ArchiveRecord(13, catalogType, Message(includeCatalogKind ? VarintField(1, field == 4 ? 1u : 3u) : Message(), payload)));
    }
}
