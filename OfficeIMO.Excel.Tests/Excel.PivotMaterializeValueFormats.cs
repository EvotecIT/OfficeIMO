using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void MaterializePivotTable_FormatsValuesWithTheDataFieldNumberFormat() {
        string path = Path.Combine(_directoryWithFiles, "PivotMaterializeValueFormats.xlsx");
        using ExcelDocument document = ExcelDocument.Create(path);
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Region");
        sheet.Cell(1, 2).SetValue("Sales");
        sheet.Cell(2, 1).SetValue("North");
        sheet.Cell(2, 2).SetValue(836000d);
        sheet.Cell(3, 1).SetValue("South");
        sheet.Cell(3, 2).SetValue(692000d);
        sheet.AddPivotTable(sourceRange: "A1:B3", destinationCell: "D1", new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, "Total Sales", numberFormat: "#,##0") } });

        sheet.MaterializePivotTable("SalesPivot");

        int formatted = 0;
        for (int row = 1; row <= 6; row++) {
            for (int column = 4; column <= 6; column++) {
                if (sheet.TryGetCellValueSnapshot(row, column, out ExcelNativeCellValueInfo? value) && value?.Kind == ExcelCellValueKind.Number) {
                    Assert.Equal("#,##0", sheet.Cell(row, column).GetStyle().NumberFormatCode);
                    formatted++;
                }
            }
        }

        Assert.True(formatted >= 3, "Expected the two region totals and the grand total to be formatted.");
    }
}
