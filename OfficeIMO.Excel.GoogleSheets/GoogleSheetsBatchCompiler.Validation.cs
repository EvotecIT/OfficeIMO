using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsBatchCompiler {
        private static void AppendValidationRanges(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            GoogleSheetsBatch batch,
            TranslationReport report,
            ref bool cellValidationNoticeAdded) {
            foreach (var validation in worksheet.Validations) {
                if (!IsSupportedDirectCellValidation(validation)) {
                    continue;
                }

                foreach (var a1Range in validation.A1Ranges) {
                    if (!TryParseValidationRange(a1Range, out int startRow, out int startColumn, out int endRow, out int endColumn)) {
                        continue;
                    }

                    var validationRule = BuildDirectCellValidationRule(workbookSnapshot, worksheet, validation, startRow, startColumn);
                    if (validationRule == null) {
                        continue;
                    }

                    if (!cellValidationNoticeAdded) {
                        report.Add(
                            TranslationSeverity.Info,
                            "CellValidations",
                            "List, whole-number, decimal, date, and text-length Excel data validations compile into native Google Sheets range validation rules without materializing empty target cells.");
                        cellValidationNoticeAdded = true;
                    }

                    batch.Add(new GoogleSheetsSetDataValidationRequest {
                        SheetName = worksheet.Name,
                        A1Range = a1Range,
                        StartRowIndex = startRow - 1,
                        EndRowIndexExclusive = endRow,
                        StartColumnIndex = startColumn - 1,
                        EndColumnIndexExclusive = endColumn,
                        Rule = validationRule,
                    });
                }
            }
        }

        private static bool TryParseValidationRange(
            string? a1Range,
            out int startRow,
            out int startColumn,
            out int endRow,
            out int endColumn) {
            startRow = startColumn = endRow = endColumn = 0;
            if (string.IsNullOrWhiteSpace(a1Range)) {
                return false;
            }

            var normalizedRange = a1Range!.Replace("$", string.Empty);
            if (!A1.TryParseRange(normalizedRange, out startRow, out startColumn, out endRow, out endColumn)) {
                var (singleRow, singleColumn) = A1.ParseCellRef(normalizedRange);
                startRow = endRow = singleRow;
                startColumn = endColumn = singleColumn;
            }

            return startRow > 0 && startColumn > 0 && endRow >= startRow && endColumn >= startColumn;
        }

        private static bool IsSupportedDirectCellValidation(ExcelDataValidationSnapshot validation) {
            return string.Equals(validation.Type, "list", StringComparison.OrdinalIgnoreCase)
                || string.Equals(validation.Type, "whole", StringComparison.OrdinalIgnoreCase)
                || string.Equals(validation.Type, "decimal", StringComparison.OrdinalIgnoreCase)
                || string.Equals(validation.Type, "date", StringComparison.OrdinalIgnoreCase)
                || string.Equals(validation.Type, "textlength", StringComparison.OrdinalIgnoreCase);
        }

        private static GoogleSheetsDataValidationRule? BuildCellValidationRule(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            int row,
            int column,
            TranslationReport report,
            ref bool cellValidationNoticeAdded) {
            foreach (var validation in worksheet.Validations) {
                if (!ValidationAppliesToCell(validation, row, column)) {
                    continue;
                }

                var rule = BuildDirectCellValidationRule(workbookSnapshot, worksheet, validation, row, column);
                if (rule == null) {
                    continue;
                }

                if (!cellValidationNoticeAdded) {
                    report.Add(
                        TranslationSeverity.Info,
                        "CellValidations",
                        "List, whole-number, decimal, date, and text-length Excel data validations now compile into native Google Sheets cell validation rules for populated and empty target cells within explicit ranges.");
                    cellValidationNoticeAdded = true;
                }

                return rule;
            }

            return null;
        }

        private static GoogleSheetsDataValidationRule? BuildDirectCellValidationRule(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            ExcelDataValidationSnapshot validation,
            int row,
            int column) {
            if (string.Equals(validation.Type, "list", StringComparison.OrdinalIgnoreCase)) {
                if (IsListValidationHandledByNativeTable(worksheet, validation, row, column)) {
                    return null;
                }

                var listValues = ResolveListValidationValues(workbookSnapshot, worksheet, validation);
                if (listValues.Count == 0) {
                    return null;
                }

                return new GoogleSheetsDataValidationRule {
                    ConditionType = "ONE_OF_LIST",
                    Values = listValues,
                    Strict = true,
                    ShowCustomUi = true,
                };
            }

            if (string.Equals(validation.Type, "whole", StringComparison.OrdinalIgnoreCase)
                || string.Equals(validation.Type, "decimal", StringComparison.OrdinalIgnoreCase)) {
                if (!TryMapNumericValidationConditionType(validation.Operator, out var numericConditionType, out var numericRequiresSecondValue)) {
                    return null;
                }

                if (!TryParseValidationNumber(validation.Formula1, out var firstNumberValue)) {
                    return null;
                }

                var numericValues = new List<string> { firstNumberValue };
                if (numericRequiresSecondValue) {
                    if (!TryParseValidationNumber(validation.Formula2, out var secondNumberValue)) {
                        return null;
                    }

                    numericValues.Add(secondNumberValue);
                }

                return new GoogleSheetsDataValidationRule {
                    ConditionType = numericConditionType,
                    Values = numericValues,
                    Strict = true,
                    ShowCustomUi = false,
                };
            }

            if (string.Equals(validation.Type, "date", StringComparison.OrdinalIgnoreCase)) {
                if (!TryMapDateValidationConditionType(validation.Operator, out var dateConditionType, out var dateRequiresSecondValue)) {
                    return null;
                }

                if (!TryParseValidationDate(validation.Formula1, workbookSnapshot.DateSystem, out var firstDateValue)) {
                    return null;
                }

                var dateValues = new List<string> { firstDateValue };
                if (dateRequiresSecondValue) {
                    if (!TryParseValidationDate(validation.Formula2, workbookSnapshot.DateSystem, out var secondDateValue)) {
                        return null;
                    }

                    dateValues.Add(secondDateValue);
                }

                return new GoogleSheetsDataValidationRule {
                    ConditionType = dateConditionType,
                    Values = dateValues,
                    Strict = true,
                    ShowCustomUi = false,
                };
            }

            if (string.Equals(validation.Type, "textlength", StringComparison.OrdinalIgnoreCase)) {
                if (!TryBuildTextLengthValidationFormula(validation, row, column, out var textLengthFormula)) {
                    return null;
                }

                return new GoogleSheetsDataValidationRule {
                    ConditionType = "CUSTOM_FORMULA",
                    Values = new[] { textLengthFormula },
                    Strict = true,
                    ShowCustomUi = false,
                };
            }

            return null;
        }

        private static bool IsListValidationHandledByNativeTable(
            ExcelWorksheetSnapshot worksheet,
            ExcelDataValidationSnapshot validation,
            int row,
            int column) {
            foreach (var table in worksheet.Tables) {
                var firstDataRow = GetFirstDataRowIndex(table);
                var lastDataRow = GetLastDataRowIndex(table);
                if (!firstDataRow.HasValue || !lastDataRow.HasValue) {
                    continue;
                }

                if (row < firstDataRow.Value || row > lastDataRow.Value) {
                    continue;
                }

                if (column < table.StartColumn || column > table.EndColumn) {
                    continue;
                }

                var matchingValidation = FindMatchingListValidation(worksheet, table, column);
                if (ReferenceEquals(matchingValidation, validation)) {
                    return true;
                }
            }

            return false;
        }

        private static IReadOnlyList<string> ResolveListValidationValues(
            ExcelWorkbookSnapshot workbookSnapshot,
            ExcelWorksheetSnapshot worksheet,
            ExcelDataValidationSnapshot validation) {
            var explicitValues = ParseExplicitListValidationValues(validation.Formula1);
            if (explicitValues.Count > 0) {
                return explicitValues;
            }

            var referencedRangeValues = ResolveReferencedRangeValidationValues(workbookSnapshot, worksheet.Name, validation.Formula1);
            if (referencedRangeValues.Count > 0) {
                return referencedRangeValues;
            }

            return ResolveNamedRangeValidationValues(workbookSnapshot, worksheet.Name, validation.Formula1);
        }

        private static ExcelDataValidationSnapshot? FindMatchingListValidation(
            ExcelWorksheetSnapshot worksheet,
            ExcelTableSnapshot table,
            int absoluteColumn) {
            var firstDataRow = GetFirstDataRowIndex(table);
            var lastDataRow = GetLastDataRowIndex(table);
            if (!firstDataRow.HasValue || !lastDataRow.HasValue) {
                return null;
            }

            var columnAddress = A1.ColumnIndexToLetters(absoluteColumn);
            var expectedRange = $"{columnAddress}{firstDataRow.Value}:{columnAddress}{lastDataRow.Value}";
            return worksheet.Validations.FirstOrDefault(validation =>
                string.Equals(validation.Type, "list", StringComparison.OrdinalIgnoreCase)
                && validation.A1Ranges.Count == 1
                && string.Equals(validation.A1Ranges[0], expectedRange, StringComparison.OrdinalIgnoreCase));
        }

        private static bool ValidationAppliesToCell(ExcelDataValidationSnapshot validation, int row, int column) {
            foreach (var a1Range in validation.A1Ranges) {
                if (TryRangeContainsCell(a1Range, row, column)) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryRangeContainsCell(string a1Range, int row, int column) {
            if (string.IsNullOrWhiteSpace(a1Range)) {
                return false;
            }

            var normalizedRange = a1Range.Replace("$", string.Empty);
            int startRow;
            int startColumn;
            int endRow;
            int endColumn;

            if (!A1.TryParseRange(normalizedRange, out startRow, out startColumn, out endRow, out endColumn)) {
                var (singleRow, singleColumn) = A1.ParseCellRef(normalizedRange);
                if (singleRow <= 0 || singleColumn <= 0) {
                    return false;
                }

                startRow = endRow = singleRow;
                startColumn = endColumn = singleColumn;
            }

            return row >= startRow
                && row <= endRow
                && column >= startColumn
                && column <= endColumn;
        }

        private static bool TryMapNumericValidationConditionType(
            string? validationOperator,
            out string conditionType,
            out bool requiresSecondValue) {
            requiresSecondValue = false;

            switch (validationOperator) {
                case "between":
                    conditionType = "NUMBER_BETWEEN";
                    requiresSecondValue = true;
                    return true;
                case "notBetween":
                    conditionType = "NUMBER_NOT_BETWEEN";
                    requiresSecondValue = true;
                    return true;
                case "equal":
                    conditionType = "NUMBER_EQ";
                    return true;
                case "notEqual":
                    conditionType = "NUMBER_NOT_EQ";
                    return true;
                case "greaterThan":
                    conditionType = "NUMBER_GREATER";
                    return true;
                case "greaterThanOrEqual":
                    conditionType = "NUMBER_GREATER_THAN_EQ";
                    return true;
                case "lessThan":
                    conditionType = "NUMBER_LESS";
                    return true;
                case "lessThanOrEqual":
                    conditionType = "NUMBER_LESS_THAN_EQ";
                    return true;
                default:
                    conditionType = string.Empty;
                    return false;
            }
        }

        private static bool TryMapDateValidationConditionType(
            string? validationOperator,
            out string conditionType,
            out bool requiresSecondValue) {
            requiresSecondValue = false;

            switch (validationOperator) {
                case "between":
                    conditionType = "DATE_BETWEEN";
                    requiresSecondValue = true;
                    return true;
                case "notBetween":
                    conditionType = "DATE_NOT_BETWEEN";
                    requiresSecondValue = true;
                    return true;
                case "equal":
                    conditionType = "DATE_EQ";
                    return true;
                case "greaterThan":
                    conditionType = "DATE_AFTER";
                    return true;
                case "greaterThanOrEqual":
                    conditionType = "DATE_ON_OR_AFTER";
                    return true;
                case "lessThan":
                    conditionType = "DATE_BEFORE";
                    return true;
                case "lessThanOrEqual":
                    conditionType = "DATE_ON_OR_BEFORE";
                    return true;
                default:
                    conditionType = string.Empty;
                    return false;
            }
        }

        private static bool TryParseValidationNumber(string? value, out string normalizedNumber) {
            normalizedNumber = string.Empty;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            if (!double.TryParse(value, NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var parsed)) {
                return false;
            }

            normalizedNumber = parsed.ToString("G15", CultureInfo.InvariantCulture);
            return true;
        }

        private static bool TryParseValidationDate(
            string? value,
            ExcelDateSystem dateSystem,
            out string normalizedDate) {
            normalizedDate = string.Empty;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            if (!double.TryParse(value, NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var serialDate)) {
                return false;
            }

            try {
                normalizedDate = ExcelDateSystemConverter.FromSerial(serialDate, dateSystem)
                    .ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

        private static bool TryBuildTextLengthValidationFormula(
            ExcelDataValidationSnapshot validation,
            int row,
            int column,
            out string formula) {
            formula = string.Empty;

            if (row <= 0 || column <= 0) {
                return false;
            }

            var cellReference = A1.ColumnIndexToLetters(column) + row.ToString(CultureInfo.InvariantCulture);
            var lengthExpression = $"LEN({cellReference})";

            if (!TryParseValidationNumber(validation.Formula1, out var firstValue)) {
                return false;
            }

            switch (validation.Operator) {
                case "equal":
                    formula = $"={lengthExpression}={firstValue}";
                    return true;
                case "notEqual":
                    formula = $"={lengthExpression}<>{firstValue}";
                    return true;
                case "greaterThan":
                    formula = $"={lengthExpression}>{firstValue}";
                    return true;
                case "greaterThanOrEqual":
                    formula = $"={lengthExpression}>={firstValue}";
                    return true;
                case "lessThan":
                    formula = $"={lengthExpression}<{firstValue}";
                    return true;
                case "lessThanOrEqual":
                    formula = $"={lengthExpression}<={firstValue}";
                    return true;
                case "between":
                    if (!TryParseValidationNumber(validation.Formula2, out var secondBetweenValue)) {
                        return false;
                    }

                    formula = $"=AND({lengthExpression}>={firstValue},{lengthExpression}<={secondBetweenValue})";
                    return true;
                case "notBetween":
                    if (!TryParseValidationNumber(validation.Formula2, out var secondNotBetweenValue)) {
                        return false;
                    }

                    formula = $"=OR({lengthExpression}<{firstValue},{lengthExpression}>{secondNotBetweenValue})";
                    return true;
                default:
                    return false;
            }
        }

        private static IReadOnlyList<string> ResolveReferencedRangeValidationValues(
            ExcelWorkbookSnapshot workbookSnapshot,
            string sourceSheetName,
            string? formula1) {
            if (string.IsNullOrWhiteSpace(formula1)) {
                return Array.Empty<string>();
            }

            var referenceText = formula1!.Trim();
            if (referenceText.StartsWith("=", StringComparison.Ordinal)) {
                referenceText = referenceText.Substring(1).Trim();
            }

            if (string.IsNullOrWhiteSpace(referenceText)) {
                return Array.Empty<string>();
            }

            var targetSheetName = sourceSheetName;
            if (TrySplitSheetQualifiedRange(referenceText, out var explicitSheetName, out var unqualifiedRange)) {
                targetSheetName = explicitSheetName!;
                referenceText = unqualifiedRange;
            }

            return ResolveWorksheetRangeValidationValues(workbookSnapshot, targetSheetName, referenceText);
        }

        private static IReadOnlyList<string> ResolveNamedRangeValidationValues(
            ExcelWorkbookSnapshot workbookSnapshot,
            string sourceSheetName,
            string? formula1) {
            if (string.IsNullOrWhiteSpace(formula1)) {
                return Array.Empty<string>();
            }

            var namedRangeName = formula1!.Trim();
            if (namedRangeName.StartsWith("=", StringComparison.Ordinal)) {
                namedRangeName = namedRangeName.Substring(1).Trim();
            }

            if (string.IsNullOrWhiteSpace(namedRangeName)) {
                return Array.Empty<string>();
            }

            var namedRange = workbookSnapshot.NamedRanges.FirstOrDefault(range =>
                string.Equals(range.Name, namedRangeName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(range.SheetName, sourceSheetName, StringComparison.OrdinalIgnoreCase))
                ?? workbookSnapshot.NamedRanges.FirstOrDefault(range =>
                    string.Equals(range.Name, namedRangeName, StringComparison.OrdinalIgnoreCase)
                    && string.IsNullOrWhiteSpace(range.SheetName));

            if (namedRange == null) {
                return Array.Empty<string>();
            }

            var rangeText = namedRange.ReferenceA1.Replace("$", string.Empty);
            var targetSheetName = namedRange.SheetName;
            if (TrySplitSheetQualifiedRange(rangeText, out var explicitSheetName, out var unqualifiedRange)) {
                targetSheetName = explicitSheetName;
                rangeText = unqualifiedRange;
            }

            if (string.IsNullOrWhiteSpace(targetSheetName)) {
                return Array.Empty<string>();
            }

            return ResolveWorksheetRangeValidationValues(workbookSnapshot, targetSheetName!, rangeText);
        }

        private static IReadOnlyList<string> ResolveWorksheetRangeValidationValues(
            ExcelWorkbookSnapshot workbookSnapshot,
            string targetSheetName,
            string rangeText) {
            if (string.IsNullOrWhiteSpace(targetSheetName) || string.IsNullOrWhiteSpace(rangeText)) {
                return Array.Empty<string>();
            }

            var targetWorksheet = workbookSnapshot.Worksheets.FirstOrDefault(worksheet =>
                string.Equals(worksheet.Name, targetSheetName, StringComparison.OrdinalIgnoreCase));
            if (targetWorksheet == null) {
                return Array.Empty<string>();
            }

            int startRow;
            int startColumn;
            int endRow;
            int endColumn;
            if (!A1.TryParseRange(rangeText, out startRow, out startColumn, out endRow, out endColumn)) {
                var (row, column) = A1.ParseCellRef(rangeText);
                if (row <= 0 || column <= 0) {
                    return Array.Empty<string>();
                }

                startRow = endRow = row;
                startColumn = endColumn = column;
            }

            if (startRow != endRow && startColumn != endColumn) {
                return Array.Empty<string>();
            }

            return targetWorksheet.Cells
                .Where(cell => cell.Row >= startRow
                    && cell.Row <= endRow
                    && cell.Column >= startColumn
                    && cell.Column <= endColumn)
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .Select(cell => ConvertCellValueToValidationItem(cell.Value))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!)
                .ToList();
        }

        private static IReadOnlyList<string> ParseExplicitListValidationValues(string? formula1) {
            if (string.IsNullOrWhiteSpace(formula1) || formula1!.Length < 2 || formula1[0] != '"' || formula1[formula1.Length - 1] != '"') {
                return Array.Empty<string>();
            }

            var values = new List<string>();
            var current = new System.Text.StringBuilder();
            var inner = formula1.Substring(1, formula1.Length - 2);

            for (int index = 0; index < inner.Length; index++) {
                var character = inner[index];
                if (character == '"'
                    && index + 1 < inner.Length
                    && inner[index + 1] == '"') {
                    current.Append('"');
                    index++;
                    continue;
                }

                if (character == ',') {
                    values.Add(current.ToString());
                    current.Clear();
                    continue;
                }

                current.Append(character);
            }

            values.Add(current.ToString());
            return values
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .ToList();
        }

        private static string? ConvertCellValueToValidationItem(object? value) {
            if (value == null) {
                return null;
            }

            return value switch {
                string text => text,
                bool boolean => boolean ? "TRUE" : "FALSE",
                DateTime dateTime => dateTime.ToString("O", CultureInfo.InvariantCulture),
                DateTimeOffset dateTimeOffset => dateTimeOffset.ToString("O", CultureInfo.InvariantCulture),
                IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
                _ => Convert.ToString(value, CultureInfo.InvariantCulture),
            };
        }

        private static bool TrySplitSheetQualifiedRange(string value, out string? sheetName, out string unqualifiedRange) {
            sheetName = null;
            unqualifiedRange = value;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            var bangIndex = value.LastIndexOf('!');
            if (bangIndex <= 0 || bangIndex >= value.Length - 1) {
                return false;
            }

            var sheetPart = value.Substring(0, bangIndex).Trim();
            var rangePart = value.Substring(bangIndex + 1).Trim();
            if (sheetPart.Length >= 2 && sheetPart[0] == '\'' && sheetPart[sheetPart.Length - 1] == '\'') {
                sheetPart = sheetPart.Substring(1, sheetPart.Length - 2).Replace("''", "'");
            }

            sheetName = sheetPart;
            unqualifiedRange = rangePart.Replace("$", string.Empty);
            return true;
        }
    }
}
