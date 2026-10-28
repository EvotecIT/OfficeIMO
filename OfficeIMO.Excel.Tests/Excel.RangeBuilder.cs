using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public class ExcelRangeBuilderTests
    {
        private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference != null && c.CellReference.Value == cellReference);
            if (cell.DataType?.Value == CellValues.InlineString) {
                return cell.InlineString?.InnerText ?? string.Empty;
            }
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
        public void RangeBuilderSetsValues()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            object[, ] values =
            {
                {
                    "A",
                    "B"
                },
                {
                    "C",
                    "D"
                }
            };
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet1 = document.AddSheet("Data");

                    var range3 = sheet1.Range("A1:B2");
                    range3.SetValues(values);
                }

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath))
            {
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
                Assert.Equal("A", GetCellValue(document._spreadSheetDocument, sheetPart, "A1"));
                Assert.Equal("D", GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RangeBuilderSetsSingleCell()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet4 = document.AddSheet("Data");

                    var range6 = sheet4.Range("A1:C3");
                    sheet4.Cell(range6.CellAt(2, 2).Row, range6.CellAt(2, 2).Column, "X");
                }

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath))
            {
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
                Assert.Equal("X", GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RangeBuilderAppliesNumberFormat()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            object[, ] values =
            {
                {
                    1.2,
                    3.4
                },
                {
                    5.6,
                    7.8
                }
            };
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet7 = document.AddSheet("Data");

                    var range9 = sheet7.Range("A1:B2");
                    range9.SetValues(values);
                    range9.SetNumberFormat("0.00");
                }

                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = spreadsheet.WorkbookPart;
                Assert.NotNull(workbookPart);
                var wsPart = workbookPart.WorksheetParts.First();
                var cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
                Assert.NotNull(cell.StyleIndex);
                uint styleIndex = cell.StyleIndex.Value;
                var stylesPart = workbookPart.WorkbookStylesPart;
                Assert.NotNull(stylesPart);
                var styles = stylesPart.Stylesheet;
                Assert.NotNull(styles);
                Assert.NotNull(styles.CellFormats);
                var cellFormat = (CellFormat)styles.CellFormats.ElementAt((int)styleIndex);
                Assert.NotNull(cellFormat.NumberFormatId);
                var nfId = cellFormat.NumberFormatId.Value;
                var numberingFormats = styles.NumberingFormats;
                Assert.NotNull(numberingFormats);
                var numberingFormat = numberingFormats.Elements<NumberingFormat>().First(n => n.NumberFormatId != null && n.NumberFormatId.Value == nfId);
                Assert.NotNull(numberingFormat.FormatCode);
                Assert.Equal("0.00", numberingFormat.FormatCode.Value);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RangeBuilderCanClearValues()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            object[, ] values =
            {
                {
                    "A",
                    "B"
                },
                {
                    "C",
                    "D"
                }
            };
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet10 = document.AddSheet("Data");

                    var range12 = sheet10.Range("A1:B2");
                    range12.SetValues(values);
                    for (int r = range12.FirstRow; r <= range12.LastRow; r++)
                        for (int column = range12.FirstColumn; column <= range12.LastColumn; column++)
                            sheet10.CellValue(r, column, string.Empty);
                }

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath))
            {
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
                Assert.Equal(string.Empty, GetCellValue(document._spreadSheetDocument, sheetPart, "A1"));
                Assert.Equal(string.Empty, GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
            }

            File.Delete(filePath);
        }
    }
}