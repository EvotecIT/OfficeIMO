using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsBatchCompiler {
        internal static IReadOnlyList<GoogleSheetsRequest> BuildFilterRequests(
            ExcelWorksheetSnapshot worksheet,
            TranslationReport report,
            ref bool multipleFilterNoticeAdded,
            ref bool customFilterNoticeAdded) {
            var requests = new List<GoogleSheetsRequest>();
            var filterSources = new List<(ExcelAutoFilterSnapshot Filter, string Title)>();

            if (worksheet.AutoFilter != null) {
                filterSources.Add((worksheet.AutoFilter, worksheet.Name + " Filter"));
            }

            foreach (var table in worksheet.Tables) {
                if (table.AutoFilter != null) {
                    var title = string.IsNullOrWhiteSpace(table.Name)
                        ? worksheet.Name + " Table Filter"
                        : table.Name + " Filter";
                    filterSources.Add((table.AutoFilter, title));
                }
            }

            if (filterSources.Count == 0) {
                return requests;
            }

            if (filterSources.Count > 1 && !multipleFilterNoticeAdded) {
                report.Add(
                    TranslationSeverity.Info,
                    "MultipleFilters",
                    "When multiple Excel filter ranges exist on one sheet, the first is emitted as the sheet basic filter and the rest are emitted as Google filter views.");
                multipleFilterNoticeAdded = true;
            }

            for (int i = 0; i < filterSources.Count; i++) {
                var source = filterSources[i];
                var criteria = BuildFilterCriteria(worksheet, source.Filter, report, ref customFilterNoticeAdded);
                if (i == 0) {
                    requests.Add(new GoogleSheetsSetBasicFilterRequest {
                        SheetName = worksheet.Name,
                        A1Range = source.Filter.A1Range,
                        StartRowIndex = source.Filter.StartRow - 1,
                        EndRowIndexExclusive = source.Filter.EndRow,
                        StartColumnIndex = source.Filter.StartColumn - 1,
                        EndColumnIndexExclusive = source.Filter.EndColumn,
                        Criteria = criteria,
                    });
                } else {
                    requests.Add(new GoogleSheetsAddFilterViewRequest {
                        SheetName = worksheet.Name,
                        Title = source.Title,
                        A1Range = source.Filter.A1Range,
                        StartRowIndex = source.Filter.StartRow - 1,
                        EndRowIndexExclusive = source.Filter.EndRow,
                        StartColumnIndex = source.Filter.StartColumn - 1,
                        EndColumnIndexExclusive = source.Filter.EndColumn,
                        Criteria = criteria,
                    });
                }
            }

            return requests;
        }

        private static IReadOnlyList<GoogleSheetsFilterColumnCriteria> BuildFilterCriteria(
            ExcelWorksheetSnapshot worksheet,
            ExcelAutoFilterSnapshot filter,
            TranslationReport report,
            ref bool customFilterNoticeAdded) {
            var criteria = new List<GoogleSheetsFilterColumnCriteria>();
            if (filter.Columns.Count == 0) {
                return criteria;
            }

            var cellMap = worksheet.Cells.ToDictionary(
                cell => GetWorksheetCellKey(cell.Row, cell.Column),
                cell => cell);

            foreach (var filterColumn in filter.Columns) {
                GoogleSheetsBooleanCondition? condition = null;
                if (filterColumn.CustomFilters != null) {
                    condition = BuildBooleanCondition(filterColumn.CustomFilters, report, ref customFilterNoticeAdded);
                }

                List<string> hiddenValues = new List<string>();
                if (filterColumn.Values.Count == 0) {
                } else {
                    var allowedValues = new HashSet<string>(filterColumn.Values, StringComparer.OrdinalIgnoreCase);
                    var observedValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var absoluteColumn = filter.StartColumn + filterColumn.ColumnId;

                    for (int row = filter.StartRow + 1; row <= filter.EndRow; row++) {
                        if (cellMap.TryGetValue(GetWorksheetCellKey(row, absoluteColumn), out var cell)) {
                            observedValues.Add(ConvertCellToFilterText(cell));
                        } else {
                            observedValues.Add(string.Empty);
                        }
                    }

                    hiddenValues = observedValues
                        .Where(value => !allowedValues.Contains(value))
                        .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }

                if (hiddenValues.Count == 0 && condition == null) {
                    continue;
                }

                criteria.Add(new GoogleSheetsFilterColumnCriteria {
                    ColumnId = filter.StartColumn + filterColumn.ColumnId - 1,
                    HiddenValues = hiddenValues,
                    Condition = condition,
                });
            }

            return criteria;
        }

        private static GoogleSheetsBooleanCondition? BuildBooleanCondition(
            ExcelCustomFiltersSnapshot customFilters,
            TranslationReport report,
            ref bool customFilterNoticeAdded) {
            if (customFilters.Conditions.Count == 2) {
                if (TryBuildNumericRangeCondition(customFilters, out var rangeCondition)) {
                    return rangeCondition;
                }

                AddUnsupportedCustomFilterNotice(report, ref customFilterNoticeAdded);
                return null;
            }

            if (customFilters.Conditions.Count != 1 || customFilters.MatchAll) {
                AddUnsupportedCustomFilterNotice(report, ref customFilterNoticeAdded);
                return null;
            }

            var condition = customFilters.Conditions[0];
            if (string.IsNullOrWhiteSpace(condition.Value)) {
                return null;
            }

            var value = condition.Value;
            var filterOperator = condition.Operator ?? "equal";
            if (TryBuildTextWildcardCondition(filterOperator, value, out var textCondition)) {
                return textCondition;
            }

            if (double.TryParse(value, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, System.Globalization.CultureInfo.InvariantCulture, out _)
                && TryBuildNumericCondition(filterOperator, value, out var numericCondition)) {
                return numericCondition;
            }

            if (TryBuildTextEqualityCondition(filterOperator, value, out var equalityCondition)) {
                return equalityCondition;
            }

            AddUnsupportedCustomFilterNotice(report, ref customFilterNoticeAdded);
            return null;
        }

        private static bool TryBuildNumericRangeCondition(
            ExcelCustomFiltersSnapshot customFilters,
            out GoogleSheetsBooleanCondition? condition) {
            condition = null;
            if (customFilters.Conditions.Count != 2) {
                return false;
            }

            var first = customFilters.Conditions[0];
            var second = customFilters.Conditions[1];
            if (!TryNormalizeNumericCondition(first, out var firstOperator, out var firstValue)
                || !TryNormalizeNumericCondition(second, out var secondOperator, out var secondValue)) {
                return false;
            }

            if (customFilters.MatchAll
                && TryGetBetweenBounds(firstOperator, firstValue, secondOperator, secondValue, out var lowerInclusive, out var upperInclusive)) {
                condition = new GoogleSheetsBooleanCondition {
                    Type = "NUMBER_BETWEEN",
                    Values = new[] { lowerInclusive, upperInclusive },
                };
                return true;
            }

            if (!customFilters.MatchAll
                && TryGetOutsideBounds(firstOperator, firstValue, secondOperator, secondValue, out var lowerExclusive, out var upperExclusive)) {
                condition = new GoogleSheetsBooleanCondition {
                    Type = "NUMBER_NOT_BETWEEN",
                    Values = new[] { lowerExclusive, upperExclusive },
                };
                return true;
            }

            return false;
        }

        private static bool TryBuildTextWildcardCondition(
            string filterOperator,
            string value,
            out GoogleSheetsBooleanCondition? condition) {
            condition = null;
            var normalizedOperator = filterOperator.Trim().ToLowerInvariant();
            var startsWithWildcard = value.StartsWith("*", StringComparison.Ordinal);
            var endsWithWildcard = value.EndsWith("*", StringComparison.Ordinal);
            var unwrappedValue = value.Trim('*');

            if (string.IsNullOrEmpty(unwrappedValue) || (!startsWithWildcard && !endsWithWildcard)) {
                return false;
            }

            string? conditionType = (normalizedOperator, startsWithWildcard, endsWithWildcard) switch {
                ("equal", true, true) => "TEXT_CONTAINS",
                ("equal", false, true) => "TEXT_STARTS_WITH",
                ("equal", true, false) => "TEXT_ENDS_WITH",
                ("notequal", true, true) => "TEXT_NOT_CONTAINS",
                _ => null,
            };

            if (conditionType == null) {
                return false;
            }

            condition = new GoogleSheetsBooleanCondition {
                Type = conditionType,
                Values = new[] { unwrappedValue },
            };
            return true;
        }

        private static bool TryBuildTextEqualityCondition(
            string filterOperator,
            string value,
            out GoogleSheetsBooleanCondition? condition) {
            condition = null;
            var normalizedOperator = filterOperator.Trim().ToLowerInvariant();
            string? conditionType = normalizedOperator switch {
                "equal" => "TEXT_EQ",
                "notequal" => "TEXT_NOT_EQ",
                _ => null,
            };

            if (conditionType == null) {
                return false;
            }

            condition = new GoogleSheetsBooleanCondition {
                Type = conditionType,
                Values = new[] { value },
            };
            return true;
        }

        private static bool TryBuildNumericCondition(
            string filterOperator,
            string value,
            out GoogleSheetsBooleanCondition? condition) {
            condition = null;
            var normalizedOperator = filterOperator.Trim().ToLowerInvariant();
            string? conditionType = normalizedOperator switch {
                "equal" => "NUMBER_EQ",
                "notequal" => "NUMBER_NOT_EQ",
                "greaterthan" => "NUMBER_GREATER",
                "greaterthanorequal" => "NUMBER_GREATER_THAN_EQ",
                "lessthan" => "NUMBER_LESS",
                "lessthanorequal" => "NUMBER_LESS_THAN_EQ",
                _ => null,
            };

            if (conditionType == null) {
                return false;
            }

            condition = new GoogleSheetsBooleanCondition {
                Type = conditionType,
                Values = new[] { value },
            };
            return true;
        }

        private static bool TryNormalizeNumericCondition(
            ExcelCustomFilterConditionSnapshot condition,
            out string normalizedOperator,
            out string normalizedValue) {
            normalizedOperator = string.Empty;
            normalizedValue = string.Empty;

            if (condition == null || string.IsNullOrWhiteSpace(condition.Value)) {
                return false;
            }

            normalizedOperator = (condition.Operator ?? string.Empty).Trim().ToLowerInvariant();
            if (!double.TryParse(condition.Value, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, System.Globalization.CultureInfo.InvariantCulture, out var numericValue)) {
                return false;
            }

            normalizedValue = numericValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return normalizedOperator is "greaterthan" or "greaterthanorequal" or "lessthan" or "lessthanorequal";
        }

        private static bool TryGetBetweenBounds(
            string firstOperator,
            string firstValue,
            string secondOperator,
            string secondValue,
            out string lowerInclusive,
            out string upperInclusive) {
            lowerInclusive = string.Empty;
            upperInclusive = string.Empty;

            if (TryMatchLowerUpper(firstOperator, firstValue, secondOperator, secondValue, out lowerInclusive, out upperInclusive)) {
                return true;
            }

            return TryMatchLowerUpper(secondOperator, secondValue, firstOperator, firstValue, out lowerInclusive, out upperInclusive);
        }

        private static bool TryGetOutsideBounds(
            string firstOperator,
            string firstValue,
            string secondOperator,
            string secondValue,
            out string lowerExclusive,
            out string upperExclusive) {
            lowerExclusive = string.Empty;
            upperExclusive = string.Empty;

            if (TryMatchOutsideRange(firstOperator, firstValue, secondOperator, secondValue, out lowerExclusive, out upperExclusive)) {
                return true;
            }

            return TryMatchOutsideRange(secondOperator, secondValue, firstOperator, firstValue, out lowerExclusive, out upperExclusive);
        }

        private static bool TryMatchLowerUpper(
            string lowerOperator,
            string lowerValue,
            string upperOperator,
            string upperValue,
            out string lowerInclusive,
            out string upperInclusive) {
            lowerInclusive = string.Empty;
            upperInclusive = string.Empty;

            if (lowerOperator != "greaterthanorequal" || upperOperator != "lessthanorequal") {
                return false;
            }

            if (double.Parse(lowerValue, System.Globalization.CultureInfo.InvariantCulture) > double.Parse(upperValue, System.Globalization.CultureInfo.InvariantCulture)) {
                return false;
            }

            lowerInclusive = lowerValue;
            upperInclusive = upperValue;
            return true;
        }

        private static bool TryMatchOutsideRange(
            string lowerOperator,
            string lowerValue,
            string upperOperator,
            string upperValue,
            out string lowerExclusive,
            out string upperExclusive) {
            lowerExclusive = string.Empty;
            upperExclusive = string.Empty;

            if (lowerOperator != "lessthan" || upperOperator != "greaterthan") {
                return false;
            }

            if (double.Parse(lowerValue, System.Globalization.CultureInfo.InvariantCulture) > double.Parse(upperValue, System.Globalization.CultureInfo.InvariantCulture)) {
                return false;
            }

            lowerExclusive = lowerValue;
            upperExclusive = upperValue;
            return true;
        }

        private static void AddUnsupportedCustomFilterNotice(
            TranslationReport report,
            ref bool customFilterNoticeAdded) {
            if (customFilterNoticeAdded) {
                return;
            }

            report.Add(
                TranslationSeverity.Info,
                "CustomFilters",
                "Single-condition Excel custom filters are translated into native Google filter conditions when possible. More complex custom filter combinations are currently preserved as diagnostics only.");
            customFilterNoticeAdded = true;
        }
    }
}
