using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void Sheets_ExposesAReadOnlyCollectionContract() {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddWorksheet("Data");

        IReadOnlyList<ExcelSheet> sheets = document.Sheets;

        Assert.Single(sheets);
        Assert.Equal("Data", sheets[0].Name);
    }

    [Fact]
    public void RowsAs_UsesThePopulatedRangeWhenNoRangeIsSpecified() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-RowsAs-{Guid.NewGuid():N}.xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "OrderId");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(2, 1, 42);
                sheet.CellValue(2, 2, 165258.24m);
                document.Save();
            }

            using ExcelDocument loaded = ExcelDocument.Load(path);
            TypedSalesRow row = Assert.Single(loaded.GetSheet("Data").RowsAs<TypedSalesRow>());

            Assert.Equal(42, row.OrderId);
            Assert.Equal(165258.24m, row.Amount);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void RangeLessTypedRowsObserveCancellationDuringUsedRangeDiscovery() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "OrderId");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            sheet.RowsAs<TypedSalesRow>(cancellationToken: cancellation.Token).ToArray());

        using var optionsCancellation = new CancellationTokenSource();
        optionsCancellation.Cancel();
        var options = new ExcelReadOptions { CancellationToken = optionsCancellation.Token };
        Assert.Throws<OperationCanceledException>(() =>
            sheet.RowsAs<TypedSalesRow>(options).ToArray());
    }

    [Fact]
    public void RowsAs_ReturnsAnEmptySequenceForAnEmptyWorksheet() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Empty");

        Assert.Empty(sheet.RowsAs<TypedSalesRow>());
    }

    [Fact]
    public void RowsAs_PreservesExcelAliasesAndTypeConverter() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Order Number");
        sheet.CellValue(1, 2, "Amount");
        sheet.CellValue(2, 1, "custom-value");
        sheet.CellValue(2, 2, 12.5m);
        var options = new ExcelReadOptions {
            TypeConverter = static (value, targetType, _) =>
                targetType == typeof(int) && Equals(value, "custom-value")
                    ? (true, 42)
                    : (false, null)
        };

        AliasedSalesRow row = Assert.Single(sheet.RowsAs<AliasedSalesRow>(options));

        Assert.Equal(42, row.OrderId);
        Assert.Equal(12.5m, row.Amount);
    }

    [Fact]
    public void RowsAs_StrictMappingRejectsUnmappedHeaders() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "OrderId");
        sheet.CellValue(1, 2, "Unexpected");
        sheet.CellValue(2, 1, 42);
        sheet.CellValue(2, 2, "value");

        DataMappingException exception = Assert.Throws<DataMappingException>(() =>
            sheet.RowsAs<TypedSalesRow>(new ExcelReadOptions { StrictTypedMapping = true }).ToArray());

        Assert.Contains("Unexpected", exception.Message);
    }

    [Fact]
    public void RowsAs_ExplicitMapperSupportsUsedAndSpecifiedRanges() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Order Number");
        sheet.CellValue(1, 2, "Amount");
        sheet.CellValue(2, 1, 42);
        sheet.CellValue(2, 2, 12.5m);

        TypedSalesRow usedRangeRow = Assert.Single(sheet.RowsAs<TypedSalesRow>(map => map
            .FromColumn<int>("Order Number", static (row, value) => { row.OrderId = value; return row; })
            .FromColumn<decimal>("Amount", static (row, value) => { row.Amount = value; return row; })));
        TypedSalesRow specifiedRangeRow = Assert.Single(sheet.RowsAs<TypedSalesRow>("A1:B2", map => map
            .FromColumn<int>("Order Number", static (row, value) => { row.OrderId = value; return row; })
            .FromColumn<decimal>("Amount", static (row, value) => { row.Amount = value; return row; })));

        Assert.Equal(42, usedRangeRow.OrderId);
        Assert.Equal(12.5m, usedRangeRow.Amount);
        Assert.Equal(usedRangeRow.OrderId, specifiedRangeRow.OrderId);
        Assert.Equal(usedRangeRow.Amount, specifiedRangeRow.Amount);
    }

    private sealed class TypedSalesRow {
        public int OrderId { get; set; }
        public decimal Amount { get; set; }
    }

    private sealed class AliasedSalesRow {
        [ExcelColumn("Order Number")]
        public int OrderId { get; set; }
        public decimal Amount { get; set; }
    }
}
