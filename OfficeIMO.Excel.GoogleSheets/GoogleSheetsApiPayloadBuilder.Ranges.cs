using System.Globalization;

namespace OfficeIMO.Excel.GoogleSheets {
    internal static partial class GoogleSheetsApiPayloadBuilder {
        private static bool TryBuildNamedRangePayload(
            IReadOnlyDictionary<string, int> sheetIds,
            GoogleSheetsAddNamedRangeRequest request,
            out GoogleSheetsApiNamedRangePayload payload) {
            payload = new GoogleSheetsApiNamedRangePayload {
                Name = request.Name,
                Range = new GoogleSheetsApiGridRangePayload(),
            };

            string? sheetName = request.SheetName;
            string rangeText = request.A1Range;

            if (TrySplitSheetQualifiedRange(rangeText, out var explicitSheet, out var unqualifiedRange)) {
                sheetName = explicitSheet;
                rangeText = unqualifiedRange;
            }

            if (string.IsNullOrWhiteSpace(sheetName)) {
                return false;
            }

            if (!sheetIds.TryGetValue(sheetName!, out var sheetId)) {
                return false;
            }

            int rowStart;
            int columnStart;
            int rowEnd;
            int columnEnd;
            if (!A1.TryParseRange(rangeText, out rowStart, out columnStart, out rowEnd, out columnEnd)) {
                var (row, column) = A1.ParseCellRef(rangeText);
                if (row <= 0 || column <= 0) {
                    return false;
                }

                rowStart = rowEnd = row;
                columnStart = columnEnd = column;
            }

            bool allRows = rowStart == 1 && rowEnd == A1.MaxRows;
            bool allColumns = columnStart == 1 && columnEnd == A1.MaxColumns;
            payload.Range = new GoogleSheetsApiGridRangePayload {
                SheetId = sheetId,
                StartRowIndex = allRows ? null : rowStart - 1,
                EndRowIndex = allRows ? null : rowEnd,
                StartColumnIndex = allColumns ? null : columnStart - 1,
                EndColumnIndex = allColumns ? null : columnEnd,
            };
            return true;
        }

        private static bool TryBuildGridRange(
            IReadOnlyDictionary<string, int> sheetIds,
            string sheetName,
            string rangeText,
            out GoogleSheetsApiGridRangePayload payload) {
            payload = new GoogleSheetsApiGridRangePayload();
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeText)) {
                return false;
            }

            if (!sheetIds.TryGetValue(sheetName, out var sheetId)) {
                return false;
            }

            string unqualifiedRange = rangeText;
            if (TrySplitSheetQualifiedRange(rangeText, out _, out var explicitRange)) {
                unqualifiedRange = explicitRange;
            }

            int rowStart;
            int columnStart;
            int rowEnd;
            int columnEnd;
            if (!A1.TryParseRange(unqualifiedRange, out rowStart, out columnStart, out rowEnd, out columnEnd)) {
                var (row, column) = A1.ParseCellRef(unqualifiedRange);
                if (row <= 0 || column <= 0) {
                    return false;
                }

                rowStart = rowEnd = row;
                columnStart = columnEnd = column;
            }

            payload = new GoogleSheetsApiGridRangePayload {
                SheetId = sheetId,
                StartRowIndex = rowStart - 1,
                EndRowIndex = rowEnd,
                StartColumnIndex = columnStart - 1,
                EndColumnIndex = columnEnd,
            };
            return true;
        }

        private static Dictionary<string, GoogleSheetsApiFilterCriteriaPayload> BuildFilterCriteriaMap(
            IReadOnlyList<GoogleSheetsFilterColumnCriteria> criteria) {
            var map = new Dictionary<string, GoogleSheetsApiFilterCriteriaPayload>();
            foreach (var criterion in criteria) {
                if ((criterion.HiddenValues == null || criterion.HiddenValues.Count == 0) && criterion.Condition == null) {
                    continue;
                }

                map[criterion.ColumnId.ToString(CultureInfo.InvariantCulture)] = new GoogleSheetsApiFilterCriteriaPayload {
                    HiddenValues = criterion.HiddenValues == null || criterion.HiddenValues.Count == 0 ? null : criterion.HiddenValues.ToList(),
                    Condition = criterion.Condition == null ? null : new GoogleSheetsApiBooleanConditionPayload {
                        Type = criterion.Condition.Type,
                        Values = criterion.Condition.Values.Select(value => new GoogleSheetsApiConditionValuePayload {
                            UserEnteredValue = value,
                        }).ToList(),
                    },
                };
            }

            return map;
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

        private static string BuildUpdateCellsFields(bool includeValue, bool includeFormat, bool includeNote, bool includeValidation) {
            var fields = new List<string>();
            if (includeValue) {
                fields.Add("userEnteredValue");
            }
            if (includeFormat) {
                fields.Add("userEnteredFormat");
                fields.Add("textFormatRuns");
            }
            if (includeValidation) {
                fields.Add("dataValidationRule");
            }
            if (includeNote) {
                fields.Add("note");
            }
            return string.Join(",", fields);
        }

        private static string BuildDimensionFields(GoogleSheetsUpdateDimensionPropertiesRequest request) {
            var fields = new List<string>();
            if (request.PixelSize.HasValue) {
                fields.Add("pixelSize");
            }
            if (request.Hidden) {
                fields.Add("hiddenByUser");
            }
            return fields.Count > 0 ? string.Join(",", fields) : "hiddenByUser";
        }

        private static string EscapeFormulaString(string value) {
            return (value ?? string.Empty).Replace("\"", "\"\"");
        }
    }
}
