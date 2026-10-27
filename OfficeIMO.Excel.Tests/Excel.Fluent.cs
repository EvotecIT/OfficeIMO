using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeColor = OfficeIMO.OfficeColor;
using Xunit;

namespace OfficeIMO.Tests
{
    public class ExcelFluentWorkbookTests
    {
        private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference != null && c.CellReference.Value == cellReference);
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                var table = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                if (table != null && int.TryParse(value, out int id)) {
                    return table.ChildElements[id].InnerText;
                }
            }
            return value;
        }

        [Fact]
        public void CanBuildWorkbookFluently()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet1 = document.AddWorksheet("Data");
                    int nextRow2 = 1;
                    sheet1.Cell(nextRow2, 1, "Name");
                    sheet1.Cell(nextRow2, 2, "Score");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Alice");
                    sheet1.Cell(nextRow2, 2, 93);
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Bob");
                    sheet1.Cell(nextRow2, 2, 88);
                    nextRow2++;
                    sheet1.AddTable("A1:B3", true, "Scores");
                    sheet1.AutoFitColumn(1);
                    sheet1.AutoFitColumn(2);
                    document.Save();
                }
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath))
            {
                Assert.Single(document.Sheets);
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
                Assert.Equal("Name", GetCellValue(document._spreadSheetDocument, sheetPart, "A1"));
                Assert.Equal("93", GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
                Assert.True(sheetPart.TableDefinitionParts.Any());
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanSetCellsWithVariousBuilders()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet3 = document.AddWorksheet("Data");
                    int nextRow4 = 1;
                    {
                        var cell1 = sheet3.CellAt("B2");
                        sheet3.Cell(cell1.Row, cell1.Column, "Sheet");
                    }

                    {
                        var cell2 = sheet3.CellAt("C" + nextRow4);
                        sheet3.Cell(cell2.Row, cell2.Column, "Row");
                    }

                    nextRow4++;
                    sheet3.Cell(2, 4, "Column");
                    var range5 = sheet3.Range("A1:D4");
                    sheet3.Cell(range5.CellAt(4, 4).Row, range5.CellAt(4, 4).Column, "Range");
                    document.Save();
                }
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath))
            {
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
                Assert.Equal("Sheet", GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
                Assert.Equal("Row", GetCellValue(document._spreadSheetDocument, sheetPart, "C1"));
                Assert.Equal("Column", GetCellValue(document._spreadSheetDocument, sheetPart, "D2"));
                Assert.Equal("Range", GetCellValue(document._spreadSheetDocument, sheetPart, "D4"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void FluentRangeRejectsAQualifierThatWouldRetargetAnotherSheet()
        {
            using ExcelDocument document = ExcelDocument.Create();
            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            {
                var sheet6 = document.AddWorksheet("Data");

                var range8 = sheet6.Range("'Other'!A1:B2");
            });
            Assert.Contains("must not qualify", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void FluentCellRejectsAQualifierThatWouldRetargetAnotherSheet()
        {
            using ExcelDocument document = ExcelDocument.Create();
            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            {
                var sheet9 = document.AddWorksheet("Data");

                {
                    var cell3 = sheet9.CellAt("'Other'!A1");
                    sheet9.Cell(cell3.Row, cell3.Column, "Wrong sheet");
                }
            });
            Assert.Contains("current worksheet", exception.Message, StringComparison.Ordinal);
            Assert.Empty(document.Inspect().Worksheets.Single().Cells);
        }

        [Fact]
        public void CanChangeColumnWidthAndHiddenState()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet11 = document.AddWorksheet("Data");
                    int nextRow12 = 1;
                    sheet11.Cell(nextRow12, 1, "Name");
                    sheet11.Cell(nextRow12, 2, "Score");
                    nextRow12++;
                    sheet11.SetColumnWidth(1, ExcelColumnWidth.Characters(25));
                    sheet11.SetColumnHidden(1, true);
                    sheet11.SetColumnWidth(2, ExcelColumnWidth.Characters(30));
                    document.Save();
                }
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = spreadsheet.WorkbookPart;
                Assert.NotNull(workbookPart);
                WorksheetPart wsPart = workbookPart.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.NotNull(columns);
                var col1 = columns.Elements<Column>().First(c => c.Min != null && c.Min.Value == 1 && c.Max != null && c.Max.Value == 1);
                var col2 = columns.Elements<Column>().First(c => c.Min != null && c.Min.Value == 2 && c.Max != null && c.Max.Value == 2);
                Assert.NotNull(col1.Width);
                Assert.Equal(25D, col1.Width.Value);
                Assert.NotNull(col1.Hidden);
                Assert.True(col1.Hidden.Value);
                Assert.NotNull(col2.Width);
                Assert.Equal(30D, col2.Width.Value);
                Assert.False(col2.Hidden?.Value ?? false);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanApplyAdvancedFeaturesFluently()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            var criteria = new Dictionary<uint, IEnumerable<string>>
            {
                {
                    0,
                    new[]
                    {
                        "Alice"
                    }
                }
            };
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet13 = document.AddWorksheet("Data");
                    int nextRow14 = 1;
                    sheet13.Cell(nextRow14, 1, "Name");
                    sheet13.Cell(nextRow14, 2, "Score");
                    nextRow14++;
                    sheet13.Cell(nextRow14, 1, "Alice");
                    sheet13.Cell(nextRow14, 2, 1);
                    nextRow14++;
                    sheet13.Cell(nextRow14, 1, "Bob");
                    sheet13.Cell(nextRow14, 2, 2);
                    nextRow14++;
                    sheet13.AddAutoFilter("A1:B3", criteria);
                    sheet13.AddConditionalColorScale("B2:B3", OfficeColor.Red, OfficeColor.Lime);
                    sheet13.AddConditionalDataBar("B2:B3", OfficeColor.Blue);
                    sheet13.AutoFitColumns();
                    sheet13.AutoFitRows();
                    document.Save();
                }
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                AutoFilter? autoFilter = wsPart.Worksheet.Elements<AutoFilter>().FirstOrDefault();
                Assert.NotNull(autoFilter);
                Assert.NotNull(autoFilter!.Reference);
                Assert.Equal("A1:B3", autoFilter.Reference!.Value);
                var rules = wsPart.Worksheet.Elements<ConditionalFormatting>().SelectMany(cf => cf.Elements<ConditionalFormattingRule>()).ToList();
                Assert.Contains(rules, r => r.Type?.Value == ConditionalFormatValues.ColorScale);
                Assert.Contains(rules, r => r.Type?.Value == ConditionalFormatValues.DataBar);
                var column = wsPart.Worksheet.GetFirstChild<Columns>()?.Elements<Column>().FirstOrDefault();
                Assert.True(column?.BestFit?.Value ?? false);
                var row = wsPart.Worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == 1);
                Assert.True(row?.CustomHeight?.Value ?? false);
                Assert.True(row?.Height?.Value > 0);
            }

            File.Delete(filePath);
        }
    }
}