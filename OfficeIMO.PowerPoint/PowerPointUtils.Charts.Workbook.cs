using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        internal static byte[] BuildChartWorkbook(PowerPointChartData data) {
            using MemoryStream ms = new();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new S.Workbook();

                WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
                var sheetData = new S.SheetData();

                int seriesCount = data.Series.Count;
                int categoryCount = data.Categories.Count;
                int totalColumns = seriesCount + 1;
                int totalRows = categoryCount + 1;
                string lastColumn = ColumnLetter(totalColumns);
                string dimensionRef = $"A1:{lastColumn}{totalRows}";

                wsPart.Worksheet = new S.Worksheet(
                    new S.SheetDimension { Reference = dimensionRef },
                    sheetData);

                var sharedStringsPart = wbPart.AddNewPart<SharedStringTablePart>();
                sharedStringsPart.SharedStringTable = new S.SharedStringTable();

                var stringIndex = new Dictionary<string, int>(StringComparer.Ordinal);
                int GetStringIndex(string value) {
                    if (!stringIndex.TryGetValue(value, out int idx)) {
                        idx = stringIndex.Count;
                        stringIndex[value] = idx;
                        sharedStringsPart.SharedStringTable.AppendChild(new S.SharedStringItem(new S.Text(value)));
                    }
                    return idx;
                }

                var headerRow = new S.Row { RowIndex = 1U, Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" } };
                headerRow.Append(CreateSharedStringCell("A1", GetStringIndex(" ")));
                for (int i = 0; i < seriesCount; i++) {
                    string cellRef = $"{ColumnLetter(i + 2)}1";
                    headerRow.Append(CreateSharedStringCell(cellRef, GetStringIndex(data.Series[i].Name)));
                }
                sheetData.Append(headerRow);

                for (int rowIndex = 0; rowIndex < categoryCount; rowIndex++) {
                    uint excelRow = (uint)(rowIndex + 2);
                    var row = new S.Row { RowIndex = excelRow, Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" } };
                    string category = data.Categories[rowIndex] ?? string.Empty;
                    row.Append(CreateSharedStringCell($"A{excelRow}", GetStringIndex(category)));

                    for (int seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                        string cellRef = $"{ColumnLetter(seriesIndex + 2)}{excelRow}";
                        double value = data.Series[seriesIndex].Values[rowIndex];
                        row.Append(CreateNumberCell(cellRef, value));
                    }
                    sheetData.Append(row);
                }

                var sheets = wbPart.Workbook.AppendChild(new S.Sheets());
                sheets.Append(new S.Sheet {
                    Id = wbPart.GetIdOfPart(wsPart),
                    SheetId = 1U,
                    Name = "Sheet1"
                });

                sharedStringsPart.SharedStringTable.Save();
                wsPart.Worksheet.Save();
                wbPart.Workbook.Save();
            }

            return ms.ToArray();
        }

        internal static byte[] BuildChartWorkbook(PowerPointScatterChartData data) {
            using MemoryStream ms = new();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new S.Workbook();

                WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
                var sheetData = new S.SheetData();

                int seriesCount = data.Series.Count;
                int totalColumns = seriesCount * 2;
                int maxPoints = data.Series.Max(series => series.XValues.Count);
                int totalRows = maxPoints + 1;
                string lastColumn = ColumnLetter(totalColumns);
                string dimensionRef = $"A1:{lastColumn}{totalRows}";

                wsPart.Worksheet = new S.Worksheet(
                    new S.SheetDimension { Reference = dimensionRef },
                    sheetData);

                var sharedStringsPart = wbPart.AddNewPart<SharedStringTablePart>();
                sharedStringsPart.SharedStringTable = new S.SharedStringTable();

                var stringIndex = new Dictionary<string, int>(StringComparer.Ordinal);
                int GetStringIndex(string value) {
                    if (!stringIndex.TryGetValue(value, out int idx)) {
                        idx = stringIndex.Count;
                        stringIndex[value] = idx;
                        sharedStringsPart.SharedStringTable.AppendChild(new S.SharedStringItem(new S.Text(value)));
                    }
                    return idx;
                }

                var headerRow = new S.Row { RowIndex = 1U, Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" } };
                for (int i = 0; i < seriesCount; i++) {
                    int xColumnIndex = (i * 2) + 1;
                    int yColumnIndex = xColumnIndex + 1;
                    string xHeaderRef = $"{ColumnLetter(xColumnIndex)}1";
                    string yHeaderRef = $"{ColumnLetter(yColumnIndex)}1";

                    headerRow.Append(CreateSharedStringCell(xHeaderRef, GetStringIndex($"{data.Series[i].Name} X")));
                    headerRow.Append(CreateSharedStringCell(yHeaderRef, GetStringIndex(data.Series[i].Name)));
                }
                sheetData.Append(headerRow);

                for (int pointIndex = 0; pointIndex < maxPoints; pointIndex++) {
                    uint excelRow = (uint)(pointIndex + 2);
                    var row = new S.Row { RowIndex = excelRow, Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" } };

                    for (int seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                        PowerPointScatterChartSeries series = data.Series[seriesIndex];
                        int xColumnIndex = (seriesIndex * 2) + 1;
                        int yColumnIndex = xColumnIndex + 1;

                        if (pointIndex < series.XValues.Count) {
                            row.Append(CreateNumberCell($"{ColumnLetter(xColumnIndex)}{excelRow}", series.XValues[pointIndex]));
                        }
                        if (pointIndex < series.YValues.Count) {
                            row.Append(CreateNumberCell($"{ColumnLetter(yColumnIndex)}{excelRow}", series.YValues[pointIndex]));
                        }
                    }

                    sheetData.Append(row);
                }

                var sheets = wbPart.Workbook.AppendChild(new S.Sheets());
                sheets.Append(new S.Sheet {
                    Id = wbPart.GetIdOfPart(wsPart),
                    SheetId = 1U,
                    Name = "Sheet1"
                });

                sharedStringsPart.SharedStringTable.Save();
                wsPart.Worksheet.Save();
                wbPart.Workbook.Save();
            }

            return ms.ToArray();
        }

        private static string ColumnLetter(int column) {
            int dividend = column;
            string columnName = string.Empty;
            while (dividend > 0) {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private static S.Cell CreateSharedStringCell(string cellReference, int sharedStringIndex) {
            return new S.Cell {
                CellReference = cellReference,
                DataType = S.CellValues.SharedString,
                CellValue = new S.CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture))
            };
        }

        private static S.Cell CreateNumberCell(string cellReference, double value) {
            return new S.Cell {
                CellReference = cellReference,
                CellValue = new S.CellValue(value.ToString(CultureInfo.InvariantCulture))
            };
        }
    }
}
