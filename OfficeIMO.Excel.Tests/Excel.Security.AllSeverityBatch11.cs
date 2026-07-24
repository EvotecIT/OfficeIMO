using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Reader_RejectsWorkbookBufferBeyondConfiguredInputLimit() {
            using var stream = new MemoryStream(new byte[32], writable: false);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocumentReader.Open(stream, new ExcelReadOptions { MaxInputBytes = 16 }));

            Assert.Contains("configured maximum size", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void PivotTable_RejectsSharedItemDiscoveryBeyondAbsoluteScanBudget() {
            using var stream = new MemoryStream();
            using ExcelDocument document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "First");
            sheet.CellValue(1, 2, "Second");
            sheet.CellValue(1, 3, "Value");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.AddPivotTable(
                    sourceRange: "A1:C1048576",
                    destinationCell: "E2",
                    rowFields: new[] { "First", "Second" },
                    dataFields: new[] { new ExcelPivotDataField("Value", DataConsolidateFunctionValues.Sum) }));

            Assert.Contains("shared-item discovery", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void AppendDataTable_ScansOutOfOrderRowsAndRejectsStyleOnlyTargetCells() {
            using var stream = new MemoryStream();
            using ExcelDocument document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Sales");
            sheet.InsertDataTableAsTable(CreateSalesTable(), tableName: "SalesTable");

            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            sheetData.Append(new Row { RowIndex = uint.MaxValue });
            sheetData.Append(new Row(new Cell { CellReference = "A3", StyleIndex = 0U }) { RowIndex = 3U });

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.AppendDataTableToTable(CreateAppendTable(), "SalesTable"));

            Assert.Contains("A3", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void AppendDataTable_UsesExplicitCellReferencesWhenRowMetadataIsHostile() {
            AssertAppendRejectsExplicitTargetCell(null);
            AssertAppendRejectsExplicitTargetCell(uint.MaxValue);
            AssertAppendRejectsExplicitTargetCell(99U);
        }

        [Fact]
        public void AppendDataTable_RejectsTargetRangeMetadataWithoutMaterializedCells() {
            using var stream = new MemoryStream();
            using ExcelDocument document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Sales");
            sheet.InsertDataTableAsTable(CreateSalesTable(), tableName: "SalesTable");
            sheet.WorksheetPart.Worksheet.AppendChild(new Hyperlinks(new Hyperlink { Reference = "A3" }));

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.AppendDataTableToTable(CreateAppendTable(), "SalesTable"));

            Assert.Contains("hyperlink", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Save_IgnoresOutOfRangeDeclaredRowsWhenRebuildingWorksheetDimension() {
            string filePath = Path.Combine(_directoryWithFiles, "SecurityBatch11.InvalidDimensionRow.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "safe");
                SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
                sheetData.Append(new Row(new Cell(new CellValue("attacker"))) { RowIndex = uint.MaxValue });
                document.Save();
            }

            using SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, false);
            Assert.Equal("A1", package.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetDimension>()!.Reference!.Value);
        }

        [Fact]
        public void SharedStringCleanup_ConvertsOutOfRangeIndexesWithoutExpandingTheTable() {
            using var stream = new MemoryStream();
            using ExcelDocument document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "safe");
            Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single();
            cell.DataType = CellValues.SharedString;
            cell.CellValue = new CellValue(int.MaxValue.ToString(System.Globalization.CultureInfo.InvariantCulture));

            document.CleanupStyleAndSharedStringArtifacts(save: false);

            Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
            SharedStringTable sharedStrings = document.WorkbookPartRoot.SharedStringTablePart!.SharedStringTable!;
            Assert.Single(sharedStrings.Elements<SharedStringItem>());
            Assert.Equal("safe", sharedStrings.Elements<SharedStringItem>().Single().InnerText);
        }

        private static DataTable CreateSalesTable() {
            var table = new DataTable();
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("Revenue", typeof(int));
            table.Rows.Add("NA", 100);
            return table;
        }

        private static DataTable CreateAppendTable() {
            var table = new DataTable();
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("Revenue", typeof(int));
            table.Rows.Add("EU", 200);
            return table;
        }

        private static void AssertAppendRejectsExplicitTargetCell(uint? declaredRowIndex) {
            using var stream = new MemoryStream();
            using ExcelDocument document = ExcelDocument.Create(stream);
            ExcelSheet sheet = document.AddWorksheet("Sales");
            sheet.InsertDataTableAsTable(CreateSalesTable(), tableName: "SalesTable");

            var hostileRow = new Row(new Cell {
                CellReference = "A3",
                CellFormula = new CellFormula("1+1")
            });
            if (declaredRowIndex.HasValue) {
                hostileRow.RowIndex = declaredRowIndex.Value;
            }
            sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Append(hostileRow);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.AppendDataTableToTable(CreateAppendTable(), "SalesTable"));

            Assert.Contains("A3", exception.Message, StringComparison.Ordinal);
        }
    }
}
