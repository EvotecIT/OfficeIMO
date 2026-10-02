using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> CatalogContractKinds() {
        foreach (int field in new[] { 4, 5, 6, 17, 19, 22 })
            foreach (string defect in new[] { "missing", "wrong", "repeated", "wire" })
                yield return new object[] { field, defect };
    }

    [Theory]
    [MemberData(nameof(CatalogContractKinds))]
    public void Selected_catalog_requires_one_correct_native_list_kind(int field, string defect) {
        byte[] metadata = defect switch {
            "missing" => Message(), "wrong" => VarintField(1, 99),
            "repeated" => Message(VarintField(1, CatalogListKind(field)), VarintField(1, CatalogListKind(field))),
            _ => BytesField(1, Message())
        };
        using var package = CatalogContractPackage(field, metadata);
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        IWorkTableCell cell = Assert.Single(table.Cells);
        if (field is 4 or 17) Assert.Equal(IWorkCellKind.Error, cell.Kind);
        else Assert.Equal(42d, cell.Value);
        if (field == 6) Assert.False(cell.FormulaIsComplete);
        if (field == 5) Assert.Null(table.GetFill(1, 1));
        if (field == 19) Assert.Null(cell.Comment);
        if (field == 22) {
            Assert.Null(cell.NumberFormat);
            Assert.Equal(IWorkCellUnsupportedFeatures.NumericFormat, cell.UnsupportedFeatures);
        }
        if (field is 19 or 22) {
            var rowIssue = Assert.Single(projection.SourceDeclarationIssues, item => item.Owner.RecordIdentifier == 12);
            Assert.Equal("5[1]/6", rowIssue.FieldPath);
            Assert.Equal(IWorkSourceDeclarationIssueKind.UnsupportedField, rowIssue.Kind);
        }
        Assert.Equal(field is 19 or 22 ? 2 : 1, projection.SourceDeclarationIssues.Count);
        var issue = Assert.Single(projection.SourceDeclarationIssues, item => item.Owner.RecordIdentifier == 13);
        Assert.Equal(13ul, issue.Owner.RecordIdentifier);
        Assert.Equal("1", issue.FieldPath);
        Assert.Equal(defect == "missing" ? 0 : defect == "repeated" ? 2 : 1, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata, issue.Kind);
    }

    [Theory]
    [InlineData(4, 6005u)] [InlineData(4, 6201u)]
    [InlineData(5, 6005u)] [InlineData(5, 6201u)]
    [InlineData(6, 6005u)] [InlineData(6, 6201u)]
    [InlineData(17, 6005u)] [InlineData(17, 6201u)]
    [InlineData(19, 6005u)] [InlineData(19, 6201u)]
    [InlineData(22, 6005u)] [InlineData(22, 6201u)]
    public void Both_registered_data_list_types_preserve_selected_values(int field, uint type) {
        using var package = CatalogContractPackage(field, VarintField(1, CatalogListKind(field)), type);
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        IWorkTableCell cell = Assert.Single(table.Cells);
        Assert.Empty(projection.SourceReferenceIssues);
        Assert.Empty(projection.SourceDeclarationIssues);
        if (field == 4) Assert.Equal("First", cell.Value);
        else if (field == 17) Assert.Equal("Value", cell.RichText!.PlainText);
        else Assert.Equal(42d, cell.Value);
        if (field == 6) Assert.True(cell.FormulaIsComplete);
        if (field == 5) Assert.NotNull(table.GetFill(1, 1));
        if (field == 19) Assert.Equal("Review", cell.Comment!.Text);
        if (field == 22) Assert.NotNull(cell.NumberFormat);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 4)] [InlineData(IWorkDocumentKind.Pages, 6)]
    [InlineData(IWorkDocumentKind.Numbers, 4)] [InlineData(IWorkDocumentKind.Numbers, 6)]
    [InlineData(IWorkDocumentKind.Keynote, 4)] [InlineData(IWorkDocumentKind.Keynote, 6)]
    public void Wrong_type_string_or_formula_catalog_cannot_decode_plausible_payload(IWorkDocumentKind kind, int field) {
        using var package = SelectedCatalogPackage(kind, field, BytesField(3, CatalogTestEntry(field)), catalogType: 2021);
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        var cell = Assert.Single(table.Cells);
        if (field == 4) Assert.Equal(IWorkCellKind.Error, cell.Kind);
        else { Assert.Equal(42d, cell.Value); Assert.False(cell.FormulaIsComplete); }
        var issue = Assert.Single(issues);
        Assert.Equal(11ul, issue.Owner.RecordIdentifier);
        Assert.Equal(field == 4 ? "4/4" : "4/6", issue.FieldPath);
        Assert.Equal(13ul, issue.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
    }

    [Theory]
    [InlineData(4)] [InlineData(5)] [InlineData(6)] [InlineData(17)] [InlineData(19)] [InlineData(22)]
    public void Invalid_list_kind_still_charges_declared_catalog_entry_work(int field) {
        using var package = CatalogContractPackage(field, VarintField(1, 99), entries: 2);
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumTableCatalogEntries = 1 });
        Assert.Contains("catalog limit", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Rejected_formula_catalog_kind_preserves_cache_after_destination_save_and_reopen(IWorkDocumentKind kind) {
        using var package = SelectedCatalogPackage(kind, 6,
            Message(VarintField(1, 1), BytesField(3, CatalogTestEntry(6))), includeCatalogKind: false);
        var source = IWorkSourceDocument.Open(package, kind);
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        IWorkConversionReport report;
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(options);
            report = result.Report;
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = WordDocument.Load(saved);
            Assert.Equal("42", reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(options);
            report = result.Report;
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
            Assert.Null(reopened.Sheets[0].CellAt(1, 1).GetValue().Formula);
        } else {
            using var result = source.ToPowerPointPresentationResult(options);
            report = result.Report;
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = PowerPointPresentation.Load(saved);
            Assert.Equal("42", reopened.Slides[0].Tables.First().GetCell(0, 0).Text);
        }
        Assert.True(report.IsPartialEditableReconstruction);
        var issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(13ul, issue.Owner.RecordIdentifier);
        Assert.Equal("1", issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata, issue.Kind);
    }

    private static ulong CatalogListKind(int field) => field switch { 4 => 1, 5 => 4, 6 => 3, 17 => 8, 19 => 10, _ => 2 };

    private static MemoryStream CatalogContractPackage(int field, byte[] metadata, uint type = 6005, int entries = 1) {
        if (field == 17) return SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            catalogPayload: Message(metadata, Enumerable.Range(0, entries).Select(_ => BytesField(3, CatalogTestEntry(field))).SelectMany(b => b).ToArray()),
            catalogType: type, includeCatalogKind: false);
        if (field == 22) return NumberFormatPackage(IWorkDocumentKind.Numbers,
            catalog: Message(metadata, Enumerable.Range(0, entries).Select(_ => BytesField(3, FormatEntry(NumericFormat(258, 2)))).SelectMany(b => b).ToArray()),
            value: 42d, catalogType: type);
        if (field is 4 or 6) return SelectedCatalogPackage(IWorkDocumentKind.Numbers, field,
            Message(metadata, Enumerable.Range(0, entries).Select(_ => BytesField(3, CatalogTestEntry(field))).SelectMany(b => b).ToArray()),
            catalogType: type, includeCatalogKind: false);
        byte[] cell = field == 19 ? CommentCell(false) : FeatureCell(false, 1u << 5);
        WriteUInt32(cell, cell.Length - 4, 1);
        byte[] entry = Message(VarintField(1, 1), ReferenceField(field == 19 ? 10 : 4, 14));
        return TableDependencyPackage(IWorkDocumentKind.Numbers, ReferenceField(field, 13), cellPayload: cell,
            additionalRecords: Message(ArchiveRecord(13, type, Message(metadata,
                Enumerable.Range(0, entries).Select(_ => BytesField(3, entry)).SelectMany(b => b).ToArray())),
                field == 19 ? ArchiveRecord(14, 3056, Message(StringField(1, "Review"), BytesField(2, DoubleField(1, 42d)), ReferenceField(3, 15)))
                    : FillStyle(14, FillColor(1, 0, 0)),
                field == 19 ? ArchiveRecord(15, 212, StringField(1, "Reviewer")) : Message()));
    }
}
