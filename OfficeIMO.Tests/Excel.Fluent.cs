using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using SixLaborsColor = SixLabors.ImageSharp.Color;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelFluentWorkbookTests {
        private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference != null && c.CellReference.Value == cellReference);
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                var table = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                if (int.TryParse(value, out int id)) {
                    return table.ChildElements[id].InnerText;
                }
            }
            return value;
        }

        [Fact]
        public void CanBuildWorkbookFluently() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => s
                        .HeaderRow("Name", "Score")
                        .Row(r => r.Values("Alice", 93))
                        .Row(r => r.Values("Bob", 88))
                        .Table(t => t.Add("A1:B3", true, "Scores"))
                        .Columns(c => c.Col(1, col => col.AutoFit()).Col(2, col => col.AutoFit())))
                    .End()
                    .Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                Assert.Single(document.Sheets);
                var sheetPart = document._spreadSheetDocument.WorkbookPart.WorksheetParts.First();
                Assert.Equal("Name", GetCellValue(document._spreadSheetDocument, sheetPart, "A1"));
                Assert.Equal("93", GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
                Assert.True(sheetPart.TableDefinitionParts.Any());
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanChangeColumnWidthAndHiddenState() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => s
                        .HeaderRow("Name", "Score")
                        .Columns(c => c
                            .Col(1, col => col.Width(25).Hidden(true))
                            .Col(2, col => col.Width(30))))
                    .End()
                    .Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.NotNull(columns);
                var col1 = columns.Elements<Column>().First(c => c.Min == 1 && c.Max == 1);
                var col2 = columns.Elements<Column>().First(c => c.Min == 2 && c.Max == 2);
                Assert.Equal(25D, col1.Width!.Value);
                Assert.True(col1.Hidden!.Value);
                Assert.Equal(30D, col2.Width!.Value);
                Assert.False(col2.Hidden?.Value ?? false);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanApplyAdvancedFeaturesFluently() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            var criteria = new Dictionary<uint, IEnumerable<string>> {
                { 0, new[] { "Alice" } }
            };

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => s
                        .HeaderRow("Name", "Score")
                        .Row(r => r.Values("Alice", 1))
                        .Row(r => r.Values("Bob", 2))
                        .AutoFilter("A1:B3", criteria)
                        .ConditionalColorScale("B2:B3", SixLaborsColor.Red, SixLaborsColor.Lime)
                        .ConditionalDataBar("B2:B3", SixLaborsColor.Blue)
                        .AutoFit(columns: true, rows: true))
                    .End()
                    .Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                AutoFilter autoFilter = wsPart.Worksheet.Elements<AutoFilter>().FirstOrDefault();
                Assert.NotNull(autoFilter);
                Assert.Equal("A1:B3", autoFilter.Reference.Value);

                var rules = wsPart.Worksheet.Elements<ConditionalFormatting>()
                    .SelectMany(cf => cf.Elements<ConditionalFormattingRule>())
                    .ToList();
                Assert.Contains(rules, r => r.Type == ConditionalFormatValues.ColorScale);
                Assert.Contains(rules, r => r.Type == ConditionalFormatValues.DataBar);

                var column = wsPart.Worksheet.GetFirstChild<Columns>()?.Elements<Column>().FirstOrDefault();
                Assert.True(column?.BestFit?.Value ?? false);

                var row = wsPart.Worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex == 1);
                Assert.True(row?.CustomHeight?.Value ?? false);
            }

            File.Delete(filePath);
        }
    }
}
