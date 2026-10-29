using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void MaterializePivotTable_FormatsValuesWithTheDataFieldNumberFormat() {
        string path = Path.Combine(_directoryWithFiles, "PivotMaterializeValueFormats.xlsx");
        using ExcelDocument document = ExcelDocument.Create(path);
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Region");
        sheet.CellValue(1, 2, "Sales");
        sheet.CellValue(2, 1, "North");
        sheet.CellValue(2, 2, 836000d);
        sheet.CellValue(3, 1, "South");
        sheet.CellValue(3, 2, 692000d);
        sheet.AddPivotTable(
            sourceRange: "A1:B3",
            destinationCell: "D1",
            name: "SalesPivot",
            rowFields: new[] { "Region" },
            dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, "Total Sales", numberFormat: "#,##0") });

        sheet.MaterializePivotTable("SalesPivot");

        int formatted = 0;
        for (int row = 1; row <= 6; row++) {
            for (int column = 4; column <= 6; column++) {
                if (sheet.TryGetCellValueSnapshot(row, column, out ExcelNativeCellValueInfo? value) && value?.Kind == ExcelCellValueKind.Number) {
                    Assert.Equal("#,##0", sheet.GetCellStyle(row, column).NumberFormatCode);
                    formatted++;
                }
            }
        }

        Assert.True(formatted >= 3, "Expected the two region totals and the grand total to be formatted.");
    }
}
