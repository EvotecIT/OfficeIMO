using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsBatchCompiler {
        private static IReadOnlyList<GoogleSheetsTableColumn> BuildTableColumns(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            var columns = new List<GoogleSheetsTableColumn>();
            foreach (var tableColumn in table.Columns) {
                var absoluteColumn = table.StartColumn + tableColumn.Index - 1;
                columns.Add(new GoogleSheetsTableColumn {
                    ColumnIndex = tableColumn.Index - 1,
                    Name = tableColumn.Name,
                    ColumnType = InferTableColumnType(workbookSnapshot, worksheet, table, absoluteColumn),
                    TotalsRowFunction = tableColumn.TotalsRowFunction,
                    DataValidationRule = BuildTableColumnValidationRule(workbookSnapshot, worksheet, table, absoluteColumn),
                });
            }

            return columns;
        }

        private static string? ResolveTableFooterColorArgb(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            if (!table.TotalsRowShown) {
                return null;
            }

            var footerColors = GetTableRowFillColors(worksheet, table, table.EndRow);

            if (footerColors.Count > 0) {
                return footerColors[0];
            }

            // A footer color is what prompts native Sheets table footer creation.
            return DefaultTableFooterColorArgb;
        }

        private static string? ResolveTableHeaderColorArgb(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            if (!table.HasHeaderRow) {
                return null;
            }

            var headerColors = GetTableRowFillColors(worksheet, table, table.StartRow);
            return headerColors.Count > 0 ? headerColors[0] : null;
        }

        private static string? ResolveTableFirstBandColorArgb(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            var firstDataRow = GetFirstDataRowIndex(table);
            if (!firstDataRow.HasValue) {
                return null;
            }

            var colors = GetTableRowFillColors(worksheet, table, firstDataRow.Value);
            return colors.Count > 0 ? colors[0] : null;
        }

        private static string? ResolveTableSecondBandColorArgb(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table) {
            var firstDataRow = GetFirstDataRowIndex(table);
            var lastDataRow = GetLastDataRowIndex(table);
            if (!firstDataRow.HasValue || !lastDataRow.HasValue) {
                return null;
            }

            var secondDataRow = firstDataRow.Value + 1;
            if (secondDataRow > lastDataRow.Value) {
                return null;
            }

            var colors = GetTableRowFillColors(worksheet, table, secondDataRow);
            return colors.Count > 0 ? colors[0] : null;
        }

        private static int? GetFirstDataRowIndex(ExcelTableSnapshot table) {
            var startRow = table.HasHeaderRow ? table.StartRow + 1 : table.StartRow;
            var lastDataRow = GetLastDataRowIndex(table);
            if (!lastDataRow.HasValue || startRow > lastDataRow.Value) {
                return null;
            }

            return startRow;
        }

        private static int? GetLastDataRowIndex(ExcelTableSnapshot table) {
            var endRow = table.TotalsRowShown ? table.EndRow - 1 : table.EndRow;
            if (endRow < table.StartRow) {
                return null;
            }

            return endRow;
        }

        private static List<string> GetTableRowFillColors(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table,
            int rowIndex) {
            return worksheet.Cells
                .Where(cell => cell.Row == rowIndex
                    && cell.Column >= table.StartColumn
                    && cell.Column <= table.EndColumn
                    && !string.IsNullOrWhiteSpace(cell.Style?.FillColorArgb))
                .Select(cell => cell.Style!.FillColorArgb!)
                .GroupBy(color => color, StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(group => group.Count())
                .ThenBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.Key)
                .ToList();
        }

        private static string InferTableColumnType(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table,
            int absoluteColumn) {
            var validationRule = BuildTableColumnValidationRule(workbookSnapshot, worksheet, table, absoluteColumn);
            if (validationRule != null) {
                return "DROPDOWN";
            }

            bool seenValue = false;
            bool allBoolean = true;
            bool allDateLike = true;
            bool allNumeric = true;
            bool anyPercent = false;
            bool anyCurrency = false;

            var startRow = table.HasHeaderRow ? table.StartRow + 1 : table.StartRow;
            var endRow = table.TotalsRowShown ? table.EndRow - 1 : table.EndRow;
            if (endRow < startRow) {
                return "TEXT";
            }

            foreach (var cell in worksheet.Cells.Where(c => c.Column == absoluteColumn && c.Row >= startRow && c.Row <= endRow)) {
                var value = cell.Value;
                if (value == null && string.IsNullOrWhiteSpace(cell.Formula)) {
                    continue;
                }

                seenValue = true;

                if (value is bool) {
                    allNumeric = false;
                    allDateLike = false;
                    continue;
                }

                allBoolean = false;

                if (value is DateTime || value is DateTimeOffset || cell.Style?.IsDateLike == true) {
                    allNumeric = false;
                    continue;
                }

                allDateLike = false;

                if (value is byte || value is sbyte || value is short || value is ushort
                    || value is int || value is uint || value is long || value is ulong
                    || value is float || value is double || value is decimal) {
                    var formatCode = cell.Style?.NumberFormatCode ?? string.Empty;
                    if (formatCode.IndexOf('%') >= 0) {
                        anyPercent = true;
                    }
                    if (formatCode.IndexOf('$') >= 0 || formatCode.IndexOf("z", StringComparison.OrdinalIgnoreCase) >= 0) {
                        anyCurrency = true;
                    }
                    continue;
                }

                allNumeric = false;
            }

            if (!seenValue) {
                return "TEXT";
            }
            if (allBoolean) {
                return "BOOLEAN";
            }
            if (allDateLike) {
                return "DATE_TIME";
            }
            if (allNumeric) {
                if (anyPercent) {
                    return "PERCENT";
                }
                if (anyCurrency) {
                    return "CURRENCY";
                }
                return "NUMBER";
            }

            return "TEXT";
        }

        private static GoogleSheetsDataValidationRule? BuildTableColumnValidationRule(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table,
            int absoluteColumn) {
            var validation = FindMatchingListValidation(worksheet, table, absoluteColumn);
            if (validation == null) {
                return null;
            }

            var values = ResolveListValidationValues(workbookSnapshot, worksheet, validation);
            if (values.Count == 0) {
                return null;
            }

            return new GoogleSheetsDataValidationRule {
                ConditionType = "ONE_OF_LIST",
                Values = values,
                Strict = true,
                ShowCustomUi = true,
            };
        }
    }
}
