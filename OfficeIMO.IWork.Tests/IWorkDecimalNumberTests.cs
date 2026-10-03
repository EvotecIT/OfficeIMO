using System.Numerics;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Approximate_decimal_values_keep_exact_source_text_when_formats_are_applied(IWorkDocumentKind kind) {
        byte[] cell = DecimalCell(BigInteger.One << 112, 0, withFormat: true);
        using MemoryStream package = TableDependencyPackage(kind, ReferenceField(22, 13), cellPayload: cell,
            additionalRecords: ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(NumericFormat(256, 2))))));
        (IWorkTable table, _) = ReadSelectedRichTable(IWorkSourceDocument.Open(package), kind);
        IWorkTableCell number = Assert.Single(table.Cells);
        Assert.Equal("5192296858534827628530496329220096E0", number.SourceNumberText);
        Assert.True(number.NumericValueIsApproximate);
        Assert.NotNull(number.NumberFormat);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMERIC_VALUE_APPROXIMATED"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void Approximate_decimal_formula_cache_keeps_its_formula_and_exact_source_text() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Decimal cache", 1, 1, 0d, hasFormula: true, completeFormula: true, decimal128HighBit: true)
        });
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        IWorkTableCell cell = Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells);
        Assert.Equal(IWorkCellKind.Formula, cell.Kind);
        Assert.True(cell.FormulaIsComplete);
        Assert.True(cell.CachedValueIsComplete);
        Assert.Equal(IWorkCellKind.Number, cell.ValueKind);
        Assert.True(cell.NumericValueIsApproximate);
        Assert.Equal("5192296858534827628530496329220096E0", cell.SourceNumberText);
        IWorkConversionReport report = projection.CreateConversionReport(IWorkProjectionKind.EditableReconstruction);
        Assert.Equal(IWorkFormulaCacheStatus.Approximate, Assert.Single(report.FormulaCells).CacheStatus);
        Assert.Equal(1, report.FormulaSummary.ApproximateCacheCount);
        Assert.Equal(0, report.FormulaSummary.CompleteCacheCount);
    }

    [Theory]
    [InlineData(0)] // Noncanonical coefficient.
    [InlineData(1)] // Overflow.
    [InlineData(2)] // Nonzero underflow.
    public void Unsupported_decimal_encodings_keep_decode_failure_instead_of_a_fabricated_value(int failure) {
        byte[] cell = DecimalCell(failure == 0 ? BigInteger.Pow(10, 34) : BigInteger.One,
            failure == 1 ? 600 : failure == 2 ? -600 : 0);
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), cellPayload: cell);
        IWorkTableCell number = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Equal(IWorkCellKind.Error, number.Kind);
        Assert.True(number.HasDecodeError);
        Assert.Null(number.Value);
        Assert.Null(number.SourceNumberText);
        Assert.False(number.NumericValueIsApproximate);
    }

    [Fact]
    public void Exact_decimal_source_text_counts_toward_the_projection_text_budget() {
        using MemoryStream package = CreateNumbersPackage(new[] { new TableSpec("X", 1, 1, 0d, decimal128HighBit: true) });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 32 });
        Assert.Throws<InvalidDataException>(() => source.ReadNumbers());
    }

    private static byte[] DecimalCell(BigInteger coefficient, int exponent, bool withFormat = false) {
        byte[] cell = new byte[withFormat ? 32 : 28]; cell[0] = 5; cell[1] = 2;
        WriteUInt32(cell, 8, 1u | (withFormat ? 1u << 13 : 0u));
        byte[] digits = coefficient.ToByteArray();
        Buffer.BlockCopy(digits, 0, cell, 12, Math.Min(14, digits.Length));
        int biased = exponent + 0x1820;
        cell[26] = (byte)(((biased & 0x7f) << 1) | (int)(coefficient >> 112));
        cell[27] = (byte)(biased >> 7);
        if (withFormat) WriteUInt32(cell, 28, 1);
        return cell;
    }
}
