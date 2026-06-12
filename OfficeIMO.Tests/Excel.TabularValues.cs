using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void InsertObjectsWithTable_UsesDirectWriterForCleanWorkbook() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(new MemoryStream(), autoSave: false)) {
                var sheet = document.AddWorkSheet("Data");
                object?[] rows = {
                    new Dictionary<string, object?> {
                        ["Region"] = "NA",
                        ["Revenue"] = 100,
                        ["Created"] = new DateTime(2026, 1, 1),
                        ["Enabled"] = true
                    },
                    new Dictionary<string, object?> {
                        ["Region"] = "EMEA",
                        ["Revenue"] = 200,
                        ["Created"] = new DateTime(2026, 1, 2),
                        ["Enabled"] = false
                    }
                };

                sheet.InsertObjects(rows);
                string range = sheet.GetUsedRangeA1();
                Assert.Equal("A1:D3", range);
                sheet.AddTable(range, hasHeader: true, name: "Sales", style: OfficeIMO.Excel.TableStyle.TableStyleMedium9, includeAutoFilter: true);

                document.Save(memory);
                Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
                Assert.True(document.LastSaveDiagnostics.UsedFastPackageWriter);
            }

            memory.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            var table = worksheetPart.TableDefinitionParts.Single().Table;
            Assert.NotNull(table);
            Assert.Equal("Sales", table!.Name?.Value);
            Assert.Equal("A1:D3", table.Reference?.Value);

            var worksheet = worksheetPart.Worksheet;
            Assert.NotNull(worksheet);
            Dictionary<string, Cell> cells = worksheet!.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("Region", GetSpreadsheetCellText(spreadsheet, cells["A1"]));
            Assert.Equal("NA", GetSpreadsheetCellText(spreadsheet, cells["A2"]));
            Assert.Equal("200", GetSpreadsheetCellText(spreadsheet, cells["B3"]));
            Assert.Equal(CellValues.Boolean, cells["D3"].DataType!.Value);
            Assert.Equal("0", cells["D3"].CellValue!.Text);
        }
    }
}
