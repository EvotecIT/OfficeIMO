using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> ScalarFormatCases() {
        foreach (var kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            foreach (int bit in new[] { 15, 16, 17, 18 }) yield return new object[] { kind, bit };
    }

    [Theory]
    [MemberData(nameof(ScalarFormatCases))]
    public void Selected_unqualified_scalar_formats_retain_values_and_require_partial_conversion(IWorkDocumentKind kind, int bit) {
        using var package = ScalarFormatPackage(kind, bit, Message(VarintField(1, 999)));
        var source = IWorkSourceDocument.Open(package, kind, new IWorkReadOptions { PreserveSourceRecords = false });
        var cell = Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells);
        Assert.False(cell.HasDecodeError);
        Assert.NotNull(cell.Value);
        Assert.Equal(bit switch {
            15 => IWorkCellUnsupportedFeatures.DateFormat, 16 => IWorkCellUnsupportedFeatures.DurationFormat,
            17 => IWorkCellUnsupportedFeatures.TextFormat, _ => IWorkCellUnsupportedFeatures.BooleanFormat
        }, cell.UnsupportedFeatures);
        foreach (bool visual in new[] { false, true }) {
            package.Position = 0;
            var report = ConvertUnitReport(package, kind, visual, new IWorkReadOptions { PreserveSourceRecords = false });
            Assert.Empty(report.PreservedRecords);
            Assert.Empty(report.SourceCellIssues);
            AssertTileDeclaration(Assert.Single(report.SourceDeclarationIssues, d => d.Owner.RecordIdentifier == 12),
                "5[1]/6", 1, IWorkSourceDeclarationIssueKind.UnsupportedField);
            Assert.Contains(report.FidelityDiagnostics, d => d.Code == "IWORK_TABLE_CELL_FEATURES_UNASSESSED"
                && d.LossKind == OfficeConversionLossKind.Unassessed);
        }
        package.Position = 0;
        var read = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build().ReadDocument(package,
            kind == IWorkDocumentKind.Pages ? "format.pages" : kind == IWorkDocumentKind.Numbers ? "format.numbers" : "format.key");
        Assert.Contains(read.Diagnostics, d => d.Code == "IWORK_TABLE_CELL_FEATURES_UNASSESSED");
        AssertScalarFormatSavedValue(source, kind, cell);
    }

    [Theory]
    [InlineData(17, 260)]
    [InlineData(18, 1)]
    public void Default_text_and_boolean_formats_keep_editable_typed_saved_output(int bit, int formatType) {
        using var package = ScalarFormatPackage(IWorkDocumentKind.Numbers, bit, VarintField(1, (ulong)formatType));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.False(result.IsVisualFallback);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.UnsupportedFeatures);
        Assert.Empty(result.Report.SourceDeclarationIssues);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        if (bit == 17) Assert.Equal("value", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
        else Assert.True(reopened.Sheets[0].CellAt(1, 1).GetValue<bool>());
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Theory]
    [InlineData(17)]
    [InlineData(18)]
    public void Unresolved_scalar_format_catalog_keeps_typed_reference_and_feature_evidence(int bit) {
        using var package = ScalarFormatPackage(IWorkDocumentKind.Numbers, bit, Message(), catalogId: 999);
        var report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false);
        var reference = Assert.Single(report.SourceReferenceIssues);
        Assert.Equal("4/22", reference.FieldPath);
        Assert.Equal(999ul, reference.TargetIdentifier);
        Assert.Contains(report.SourceDeclarationIssues, d => d.Owner.RecordIdentifier == 12 && d.FieldPath == "5[1]/6");
    }

    [Theory]
    [InlineData(17)]
    [InlineData(18)]
    public void Truncated_scalar_format_fields_do_not_select_catalogs_or_claim_feature_presence(int bit) {
        using var package = ScalarFormatPackage(IWorkDocumentKind.Numbers, bit, Message(), truncate: true);
        var source = IWorkSourceDocument.Open(package);
        var cell = Assert.Single(ReadSelectedRichTable(source, IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.True(cell.HasDecodeError);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
        package.Position = 0;
        var report = ConvertUnitReport(package, IWorkDocumentKind.Numbers);
        Assert.Single(report.SourceCellIssues);
        Assert.DoesNotContain(report.SourceDeclarationIssues, d => d.Owner.RecordIdentifier == 13);
    }

    [Theory]
    [InlineData(0)] // Extra settings, even on the default type.
    [InlineData(1)] // Ambiguous type.
    [InlineData(2)] // Wrong type wire kind.
    [InlineData(3)] // Malformed nested format.
    public void Unassessed_default_format_payloads_keep_physical_catalog_evidence(int failure) {
        byte[] format = failure switch {
            0 => Message(VarintField(1, 260), VarintField(99, 0)),
            1 => Message(VarintField(1, 260), VarintField(1, 260)),
            2 => BytesField(1, new byte[] { 1 }),
            _ => new byte[] { 0x80 }
        };
        using var package = ScalarFormatPackage(IWorkDocumentKind.Numbers, 17, format);
        var report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false);
        var issue = Assert.Single(report.SourceDeclarationIssues, d => d.Owner.RecordIdentifier == 13);
        Assert.Equal("3[1]/6", issue.FieldPath);
        Assert.Equal(failure == 3 ? IWorkSourceDeclarationIssueKind.MalformedMessage
            : IWorkSourceDeclarationIssueKind.UnsupportedField, issue.Kind);
        Assert.Empty(report.SourceCellIssues);
    }

    [Fact]
    public void Selected_scalar_formats_propagate_parser_and_shared_catalog_limits() {
        byte[] fields = Message(new[] { VarintField(1, 1) }.Concat(
            Enumerable.Range(0, 17).Select(_ => VarintField(99, 0))).ToArray());
        using var parserLimited = ScalarFormatPackage(IWorkDocumentKind.Numbers, 18, fields);
        Assert.Contains("field", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(parserLimited,
            new IWorkReadOptions { MaximumProtobufFieldCount = 16 }).ReadNumbers()).Message, StringComparison.OrdinalIgnoreCase);
        using var catalogLimited = ScalarFormatPackage(IWorkDocumentKind.Numbers, 18, VarintField(1, 1), extraFormat: true);
        Assert.Contains("catalog", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(catalogLimited,
            new IWorkReadOptions { MaximumTableCatalogEntries = 1 }).ReadNumbers()).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Unused_scalar_format_catalogs_remain_unselected() {
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, ReferenceField(22, 999));
        var projection = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumTableCatalogEntries = 1 }).ReadNumbers();
        Assert.True(projection.HasEditableContent);
        Assert.Empty(projection.SourceReferenceIssues);
        Assert.Empty(projection.SourceDeclarationIssues);
    }

    [Fact]
    public void Date_format_only_empty_cells_remain_materialized_under_the_existing_cell_budget() {
        byte[] cell = FeatureCell(empty: true, (1u << 12) | (1u << 15));
        WriteUInt32(cell, 12, 3);
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), cellPayload: cell);
        var selected = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Equal(IWorkCellKind.Empty, selected.Kind);
        Assert.Equal(IWorkCellUnsupportedFeatures.DateFormat, selected.UnsupportedFeatures);
        using var overBudget = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), columns: 2,
            tilePayload: BytesField(5, Message(VarintField(1, 0), BytesField(6, Message(cell, cell)),
                BytesField(7, new byte[] { 0, 0, (byte)cell.Length, 0 }))));
        Assert.Contains("source-wide limit", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(overBudget,
            new IWorkReadOptions { MaximumMaterializedCells = 1 }).ReadNumbers()).Message);
    }

    private static void AssertScalarFormatSavedValue(IWorkSourceDocument source, IWorkDocumentKind kind, IWorkTableCell cell) {
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var automatic = source.ToWordDocumentResult(); Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToWordDocumentResult(options); partial.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            Assert.Equal(cell.CachedDisplayText, reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Empty(reopened.ValidateDocument());
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var automatic = source.ToExcelDocumentResult(); Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToExcelDocumentResult(options); partial.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            if (cell.Value is DateTime date) Assert.Equal(date.ToOADate(), reopened.Sheets[0].CellAt(1, 1).GetValue<double>(), 10);
            else if (cell.Value is bool) Assert.Equal(cell.Value, reopened.Sheets[0].CellAt(1, 1).GetValue<bool>());
            else if (cell.Value is string) Assert.Equal(cell.Value, reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
            else Assert.Equal(cell.Kind == IWorkCellKind.Duration ? (double)cell.Value! / 86_400d : cell.Value,
                reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
            Assert.Empty(reopened.ValidateOpenXml());
        } else {
            using var automatic = source.ToPowerPointPresentationResult(); Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToPowerPointPresentationResult(options); partial.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            Assert.Equal(cell.CachedDisplayText, Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0).Text);
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    private static MemoryStream ScalarFormatPackage(IWorkDocumentKind kind, int bit, byte[] format, ulong catalogId = 13,
        bool truncate = false, bool extraFormat = false) {
        bool text = bit == 17;
        int valueBit = bit == 15 ? 2 : text ? 3 : 1;
        byte[] cell = new byte[text ? 20 : 24]; cell[0] = 5;
        cell[1] = bit == 15 ? (byte)5 : bit == 16 ? (byte)7 : text ? (byte)3 : (byte)6;
        WriteUInt32(cell, 8, (1u << valueBit) | (1u << bit));
        if (text) WriteUInt32(cell, 12, 1);
        else Buffer.BlockCopy(BitConverter.GetBytes(bit == 18 ? 1d : 42d), 0, cell, 12, 8);
        WriteUInt32(cell, cell.Length - 4, 1);
        if (truncate) cell = cell.Take(cell.Length - 1).ToArray();
        return TableDependencyPackage(kind, Message(ReferenceField(22, catalogId), text ? ReferenceField(4, 14) : Message()),
            cellPayload: cell, additionalRecords: Message(
                ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(format)),
                    extraFormat ? BytesField(3, Message(VarintField(1, 2), BytesField(6, VarintField(1, 1)))) : Message())),
                text ? ArchiveRecord(14, 6005, Message(VarintField(1, 1), BytesField(3,
                    Message(VarintField(1, 1), StringField(3, "value"))))) : Message()));
    }
}
