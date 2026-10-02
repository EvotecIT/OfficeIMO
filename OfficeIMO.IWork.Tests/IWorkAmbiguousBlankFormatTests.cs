using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 17)]
    [InlineData(IWorkDocumentKind.Numbers, 17)]
    [InlineData(IWorkDocumentKind.Keynote, 17)]
    [InlineData(IWorkDocumentKind.Numbers, 15)]
    [InlineData(IWorkDocumentKind.Numbers, 16)]
    [InlineData(IWorkDocumentKind.Numbers, 18)]
    public void Ambiguous_blank_scalar_formats_remain_visible_and_require_partial_acceptance(IWorkDocumentKind kind, int otherBit) {
        byte[] cell = new byte[24]; cell[0] = 5;
        WriteUInt32(cell, 8, (1u << 12) | (1u << 13) | (1u << otherBit));
        WriteUInt32(cell, 12, 1); WriteUInt32(cell, 16, 1); WriteUInt32(cell, 20, 2);
        byte[] other = otherBit switch {
            15 => DateFormat("dd/MM/y"),
            16 => Message(VarintField(1, 268)),
            17 => VarintField(1, 260),
            _ => VarintField(1, 1)
        };
        using var package = TableDependencyPackage(kind, ReferenceField(22, 13), cellPayload: cell,
            additionalRecords: ArchiveRecord(13, 6005, Message(VarintField(1, 2),
                BytesField(3, FormatEntry(NumericFormat(258, 2))),
                BytesField(3, Message(VarintField(1, 2), BytesField(6, other))))));
        var selected = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells);
        Assert.Equal(IWorkCellKind.Empty, selected.Kind);
        Assert.Null(selected.NumberFormat);
        Assert.True(selected.UnsupportedFeatures.HasFlag(IWorkCellUnsupportedFeatures.AmbiguousNumberFormat));
        package.Position = 0;
        var report = ConvertUnitReport(package, kind, visual: false);
        Assert.True(report.IsPartialEditableReconstruction);
        Assert.Contains(report.FidelityDiagnostics, d => d.Code == "IWORK_TABLE_CELL_FEATURES_UNASSESSED"
            && d.LossKind == OfficeConversionLossKind.Unassessed);
        Assert.Contains(report.SourceDeclarationIssues, d => d.Owner.RecordIdentifier == 12
            && d.Kind == IWorkSourceDeclarationIssueKind.UnsupportedField);
        if (kind != IWorkDocumentKind.Numbers) return;
        package.Position = 0;
        Assert.Throws<InvalidDataException>(() => ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly }));
    }
}
