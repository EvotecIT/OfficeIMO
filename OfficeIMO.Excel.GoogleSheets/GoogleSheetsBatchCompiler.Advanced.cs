using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.GoogleWorkspace;
using System.Globalization;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsBatchCompiler {
        private const string ChartDataSheetName = "_OfficeIMO_ChartData";

        private static void AppendDimensionGroups(GoogleSheetsBatch batch, ExcelWorksheetSnapshot worksheet) {
            AppendDimensionGroups(
                batch,
                worksheet.Name,
                GoogleSheetsDimensionKind.Rows,
                worksheet.Rows.Where(row => row.OutlineLevel.GetValueOrDefault() > 0)
                    .Select(row => (Start: row.Index - 1, End: row.Index, Level: row.OutlineLevel!.Value)));
            AppendDimensionGroups(
                batch,
                worksheet.Name,
                GoogleSheetsDimensionKind.Columns,
                worksheet.Columns.Where(column => column.OutlineLevel.GetValueOrDefault() > 0)
                    .Select(column => (Start: column.StartIndex - 1, End: column.EndIndex, Level: column.OutlineLevel!.Value)));
        }

        private static void AppendDimensionGroups(
            GoogleSheetsBatch batch,
            string sheetName,
            GoogleSheetsDimensionKind kind,
            IEnumerable<(int Start, int End, byte Level)> definitions) {
            var items = definitions.OrderBy(item => item.Start).ToArray();
            byte maxLevel = items.Length == 0 ? (byte)0 : items.Max(item => item.Level);
            for (byte level = 1; level <= maxLevel; level++) {
                int? start = null;
                int end = 0;
                foreach (var item in items.Where(item => item.Level >= level)) {
                    if (!start.HasValue || item.Start > end) {
                        if (start.HasValue) AddDimensionGroup(batch, sheetName, kind, start.Value, end);
                        start = item.Start;
                        end = item.End;
                    } else {
                        end = Math.Max(end, item.End);
                    }
                }
                if (start.HasValue) AddDimensionGroup(batch, sheetName, kind, start.Value, end);
            }
        }

        private static void AddDimensionGroup(GoogleSheetsBatch batch, string sheetName, GoogleSheetsDimensionKind kind, int start, int end) {
            batch.Add(new GoogleSheetsAddDimensionGroupRequest {
                SheetName = sheetName,
                DimensionKind = kind,
                StartIndex = start,
                EndIndexExclusive = end,
            });
        }

        private static void AppendConditionalFormatting(GoogleSheetsBatch batch, ExcelSheet sheet, TranslationReport report) {
            int index = 0;
            foreach (ExcelConditionalFormattingInfo rule in sheet.GetConditionalFormattingRules().OrderBy(rule => rule.Priority)) {
                if (!TryMapConditionalRule(rule, out string conditionType, out IReadOnlyList<string> values)) {
                    report.Add(
                        TranslationSeverity.Warning,
                        "ConditionalFormatting",
                        $"Conditional-formatting rule '{rule.Type}' on '{sheet.Name}!{rule.Range}' has no Google Sheets equivalent and was skipped.",
                        code: "SHEETS.CONDITIONAL_FORMAT.UNSUPPORTED",
                        action: TranslationAction.Skip,
                        targetId: sheet.Name);
                    continue;
                }

                batch.Add(new GoogleSheetsAddConditionalFormatRuleRequest {
                    SheetName = sheet.Name,
                    A1Range = rule.Range,
                    Index = index++,
                    ConditionType = conditionType,
                    Values = values,
                    Format = BuildConditionalFormat(rule),
                });
            }
        }

        private static bool TryMapConditionalRule(ExcelConditionalFormattingInfo rule, out string conditionType, out IReadOnlyList<string> values) {
            conditionType = string.Empty;
            values = rule.Formulas ?? Array.Empty<string>();
            string type = rule.Type ?? string.Empty;
            string op = rule.Operator ?? string.Empty;
            if (type.Equals("CellIs", StringComparison.OrdinalIgnoreCase)) {
                conditionType = op.ToUpperInvariant() switch {
                    "BETWEEN" => "NUMBER_BETWEEN",
                    "NOTBETWEEN" => "NUMBER_NOT_BETWEEN",
                    "EQUAL" => "NUMBER_EQ",
                    "NOTEQUAL" => "NUMBER_NOT_EQ",
                    "GREATERTHAN" => "NUMBER_GREATER",
                    "GREATERTHANOREQUAL" => "NUMBER_GREATER_THAN_EQ",
                    "LESSTHAN" => "NUMBER_LESS",
                    "LESSTHANOREQUAL" => "NUMBER_LESS_THAN_EQ",
                    _ => string.Empty,
                };
            } else if (type.Equals("Expression", StringComparison.OrdinalIgnoreCase)) {
                conditionType = "CUSTOM_FORMULA";
            } else if (type.Equals("ContainsText", StringComparison.OrdinalIgnoreCase)) {
                conditionType = "TEXT_CONTAINS";
                values = new[] { rule.Text ?? rule.Formulas?.FirstOrDefault() ?? string.Empty };
            } else if (type.Equals("NotContainsText", StringComparison.OrdinalIgnoreCase)) {
                conditionType = "TEXT_NOT_CONTAINS";
                values = new[] { rule.Text ?? rule.Formulas?.FirstOrDefault() ?? string.Empty };
            } else if (type.Equals("BeginsWith", StringComparison.OrdinalIgnoreCase)) {
                conditionType = "TEXT_STARTS_WITH";
                values = new[] { rule.Text ?? string.Empty };
            } else if (type.Equals("EndsWith", StringComparison.OrdinalIgnoreCase)) {
                conditionType = "TEXT_ENDS_WITH";
                values = new[] { rule.Text ?? string.Empty };
            } else if (type.Equals("ContainsBlanks", StringComparison.OrdinalIgnoreCase)) {
                conditionType = "BLANK";
                values = Array.Empty<string>();
            } else if (type.Equals("NotContainsBlanks", StringComparison.OrdinalIgnoreCase)) {
                conditionType = "NOT_BLANK";
                values = Array.Empty<string>();
            } else if (type.Equals("DuplicateValues", StringComparison.OrdinalIgnoreCase)) {
                conditionType = "CUSTOM_FORMULA";
                values = new[] { $"=COUNTIF({rule.Range},{FirstCell(rule.Range)})>1" };
            } else if (type.Equals("UniqueValues", StringComparison.OrdinalIgnoreCase)) {
                conditionType = "CUSTOM_FORMULA";
                values = new[] { $"=COUNTIF({rule.Range},{FirstCell(rule.Range)})=1" };
            }
            return conditionType.Length > 0;
        }

        private static string FirstCell(string range) {
            string value = range;
            int bang = value.LastIndexOf('!');
            if (bang >= 0) value = value.Substring(bang + 1);
            int colon = value.IndexOf(':');
            return (colon >= 0 ? value.Substring(0, colon) : value).Replace("$", string.Empty);
        }

        private static GoogleSheetsCellStyle BuildConditionalFormat(ExcelConditionalFormattingInfo rule) {
            return new GoogleSheetsCellStyle {
                Bold = rule.DifferentialFontBold == true,
                Italic = rule.DifferentialFontItalic == true,
                Underline = rule.DifferentialFontUnderline == true,
                FontName = rule.DifferentialFontName,
                FontSize = rule.DifferentialFontSize,
                FontColorArgb = rule.DifferentialFontColorArgb,
                FillColorArgb = rule.DifferentialFillColorArgb,
                Borders = BuildBorders(rule.DifferentialBorder),
            };
        }

        private static void AppendCharts(
            GoogleSheetsBatch batch,
            ExcelSheet sourceSheet,
            ExcelWorksheetSnapshot worksheet,
            TranslationReport report,
            UnsupportedFeatureMode policy) {
            foreach (ExcelChart chart in sourceSheet.Charts) {
                if (!chart.TryGetSnapshot(out ExcelChartSnapshot snapshot)
                    || !TryMapChartType(snapshot.ChartType, out string chartType)) {
                    HandleAdvancedUnsupported(report, "Charts", "SHEETS.CHART.UNSUPPORTED", snapshot?.Name ?? "chart", policy);
                    continue;
                }

                GoogleSheetsUpdateCellsRequest data = GetOrCreateChartDataRequest(batch, out string chartDataSheetName);
                int startRow = data.Cells.Count == 0 ? 0 : data.Cells.Max(cell => cell.RowIndex) + 2;
                data.AddCell(new GoogleSheetsCellData { RowIndex = startRow, ColumnIndex = 0, Value = GoogleSheetsCellValue.String("Category") });
                for (int seriesIndex = 0; seriesIndex < snapshot.Data.Series.Count; seriesIndex++) {
                    data.AddCell(new GoogleSheetsCellData { RowIndex = startRow, ColumnIndex = seriesIndex + 1, Value = GoogleSheetsCellValue.String(snapshot.Data.Series[seriesIndex].Name) });
                }
                for (int row = 0; row < snapshot.Data.Categories.Count; row++) {
                    data.AddCell(new GoogleSheetsCellData { RowIndex = startRow + row + 1, ColumnIndex = 0, Value = GoogleSheetsCellValue.String(snapshot.Data.Categories[row]) });
                    for (int seriesIndex = 0; seriesIndex < snapshot.Data.Series.Count; seriesIndex++) {
                        data.AddCell(new GoogleSheetsCellData { RowIndex = startRow + row + 1, ColumnIndex = seriesIndex + 1, Value = GoogleSheetsCellValue.Number(snapshot.Data.Series[seriesIndex].Values[row]) });
                    }
                }

                batch.Add(new GoogleSheetsAddChartRequest {
                    SheetName = worksheet.Name,
                    Title = snapshot.Title ?? snapshot.Name,
                    ChartType = chartType,
                    DataSheetName = chartDataSheetName,
                    DataStartRowIndex = startRow,
                    DataRowCount = snapshot.Data.Categories.Count + 1,
                    SeriesCount = snapshot.Data.Series.Count,
                    AnchorRowIndex = Math.Max(0, snapshot.RowIndex - 1),
                    AnchorColumnIndex = Math.Max(0, snapshot.ColumnIndex - 1),
                });
            }
        }

        private static GoogleSheetsUpdateCellsRequest GetOrCreateChartDataRequest(
            GoogleSheetsBatch batch,
            out string chartDataSheetName) {
            chartDataSheetName = string.IsNullOrWhiteSpace(batch.ChartDataSheetName)
                ? BuildUniqueChartDataSheetName(batch.Requests.OfType<GoogleSheetsAddSheetRequest>().Select(request => request.SheetName))
                : batch.ChartDataSheetName!;
            batch.ChartDataSheetName = chartDataSheetName;
            string existingSheetName = chartDataSheetName;
            GoogleSheetsUpdateCellsRequest? existing = batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>()
                .FirstOrDefault(request => string.Equals(request.SheetName, existingSheetName, StringComparison.OrdinalIgnoreCase));
            if (existing != null) return existing;

            int index = batch.Requests.OfType<GoogleSheetsAddSheetRequest>().Select(request => request.SheetIndex).DefaultIfEmpty(-1).Max() + 1;
            batch.Add(new GoogleSheetsAddSheetRequest { SheetName = chartDataSheetName, SheetIndex = index, Hidden = true, HideGridlines = true });
            var created = new GoogleSheetsUpdateCellsRequest { SheetName = chartDataSheetName };
            batch.Add(created);
            return created;
        }

        private static string BuildUniqueChartDataSheetName(IEnumerable<string> reservedSheetNames) {
            var reserved = new HashSet<string>(reservedSheetNames, StringComparer.OrdinalIgnoreCase);
            if (!reserved.Contains(ChartDataSheetName)) return ChartDataSheetName;

            for (int suffix = 2; ; suffix++) {
                string candidate = ChartDataSheetName + "_" + suffix.ToString(CultureInfo.InvariantCulture);
                if (!reserved.Contains(candidate)) return candidate;
            }
        }

        private static bool TryMapChartType(ExcelChartType type, out string chartType) {
            chartType = type switch {
                ExcelChartType.ColumnClustered or ExcelChartType.ColumnStacked or ExcelChartType.ColumnStacked100 => "COLUMN",
                ExcelChartType.BarClustered or ExcelChartType.BarStacked or ExcelChartType.BarStacked100 => "BAR",
                ExcelChartType.Line or ExcelChartType.LineStacked or ExcelChartType.LineStacked100 => "LINE",
                ExcelChartType.Area or ExcelChartType.AreaStacked or ExcelChartType.AreaStacked100 => "AREA",
                ExcelChartType.Pie or ExcelChartType.Doughnut => "PIE",
                ExcelChartType.Scatter => "SCATTER",
                _ => string.Empty,
            };
            return chartType.Length > 0;
        }

        private static void AppendPivotTables(
            GoogleSheetsBatch batch,
            ExcelSheet sourceSheet,
            ExcelWorkbookSnapshot workbook,
            TranslationReport report,
            UnsupportedFeatureMode policy) {
            foreach (ExcelPivotTableInfo pivot in sourceSheet.GetPivotTables()) {
                if (pivot.CalculatedFields.Count > 0 || pivot.Groupings.Count > 0 || pivot.Filters.Count > 0
                    || string.IsNullOrWhiteSpace(pivot.SourceSheet) || string.IsNullOrWhiteSpace(pivot.SourceRange)
                    || !TryGetDestination(pivot.Location, out int destinationRow, out int destinationColumn)
                    || !TryBuildPivotFields(pivot, workbook, out var rows, out var columns, out var values)) {
                    HandleAdvancedUnsupported(report, "PivotTables", "SHEETS.PIVOT_TABLE.UNSUPPORTED", pivot.Name, policy);
                    continue;
                }
                batch.Add(new GoogleSheetsAddPivotTableRequest {
                    SheetName = pivot.SheetName,
                    DestinationRowIndex = destinationRow,
                    DestinationColumnIndex = destinationColumn,
                    SourceSheetName = pivot.SourceSheet!,
                    SourceA1Range = pivot.SourceRange!,
                    Rows = rows,
                    Columns = columns,
                    Values = values,
                });
            }
        }

        private static bool TryBuildPivotFields(
            ExcelPivotTableInfo pivot,
            ExcelWorkbookSnapshot workbook,
            out IReadOnlyList<GoogleSheetsPivotGroup> rows,
            out IReadOnlyList<GoogleSheetsPivotGroup> columns,
            out IReadOnlyList<GoogleSheetsPivotValue> values) {
            rows = Array.Empty<GoogleSheetsPivotGroup>();
            columns = Array.Empty<GoogleSheetsPivotGroup>();
            values = Array.Empty<GoogleSheetsPivotValue>();
            ExcelWorksheetSnapshot? source = workbook.Worksheets.FirstOrDefault(sheet => string.Equals(sheet.Name, pivot.SourceSheet, StringComparison.OrdinalIgnoreCase));
            if (source == null || !A1.TryParseRange(pivot.SourceRange!, out int startRow, out int startColumn, out _, out _)) return false;
            var headers = source.Cells.Where(cell => cell.Row == startRow)
                .ToDictionary(cell => Convert.ToString(cell.Value) ?? string.Empty, cell => cell.Column - startColumn, StringComparer.OrdinalIgnoreCase);
            if (headers.Count == 0) return false;
            if (pivot.RowFields.Any(field => !headers.ContainsKey(field))
                || pivot.ColumnFields.Any(field => !headers.ContainsKey(field))
                || pivot.DataFields.Any(field => !headers.ContainsKey(field.FieldName))) return false;
            rows = pivot.RowFields.Select(field => new GoogleSheetsPivotGroup { SourceColumnOffset = headers[field], ShowTotals = pivot.RowGrandTotals != false }).ToArray();
            columns = pivot.ColumnFields.Select(field => new GoogleSheetsPivotGroup { SourceColumnOffset = headers[field], ShowTotals = pivot.ColumnGrandTotals != false }).ToArray();
            values = pivot.DataFields.Select(field => new GoogleSheetsPivotValue {
                SourceColumnOffset = headers[field.FieldName],
                SummarizeFunction = MapAggregate(field.Function),
                Name = field.DisplayName,
            }).ToArray();
            return values.Count > 0;
        }

        private static bool TryGetDestination(string? location, out int row, out int column) {
            row = column = 0;
            if (string.IsNullOrWhiteSpace(location)) return false;
            string first = location!.Split(':')[0].Replace("$", string.Empty);
            (int parsedRow, int parsedColumn) = A1.ParseCellRef(first);
            if (parsedRow <= 0 || parsedColumn <= 0) return false;
            row = parsedRow - 1;
            column = parsedColumn - 1;
            return true;
        }

        private static string MapAggregate(DataConsolidateFunctionValues function) {
            return function.ToString().ToUpperInvariant() switch {
                "AVERAGE" => "AVERAGE",
                "COUNT" or "COUNTNUMS" => "COUNTA",
                "MAXIMUM" => "MAX",
                "MINIMUM" => "MIN",
                "PRODUCT" => "PRODUCT",
                "STANDARDDEVIATION" or "STANDARDDEVIATIONP" => "STDEV",
                "VARIANCE" or "VARIANCEP" => "VAR",
                _ => "SUM",
            };
        }

        private static void HandleAdvancedUnsupported(TranslationReport report, string feature, string code, string target, UnsupportedFeatureMode mode) {
            TranslationSeverity severity = mode == UnsupportedFeatureMode.WarnAndSkip ? TranslationSeverity.Warning : TranslationSeverity.Error;
            report.Add(severity, feature, $"'{target}' uses {feature} features outside the native Google Sheets matrix.", code: code,
                action: mode == UnsupportedFeatureMode.WarnAndSkip ? TranslationAction.Skip : TranslationAction.Fail, targetId: target);
        }
    }
}
