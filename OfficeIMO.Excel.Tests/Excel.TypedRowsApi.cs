using System;
using System.Linq;
using System.Threading;
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
            sheet.RowsAs<TypedSalesRow>(ct: cancellation.Token).ToArray());
        Assert.Throws<OperationCanceledException>(() =>
            sheet.RowsAsStream<TypedSalesRow>(ct: cancellation.Token).ToArray());
    }

    private sealed class TypedSalesRow {
        public int OrderId { get; set; }
        public decimal Amount { get; set; }
    }
}
