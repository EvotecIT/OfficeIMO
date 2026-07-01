using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static class LegacyXlsAutoFilterProjector {
        internal static void Project(ExcelSheet sheet, string range, IReadOnlyList<LegacyXlsAutoFilterCriteria> criteria) {
            if (criteria.Count == 0 || string.IsNullOrWhiteSpace(range)) {
                return;
            }

            var equalityFilters = new Dictionary<uint, IEnumerable<string>>();
            var blankFilters = new List<LegacyXlsAutoFilterCriteria>();
            var blankOrValueFilters = new List<LegacyXlsAutoFilterCriteria>();
            var customFilters = new List<LegacyXlsAutoFilterCriteria>();
            var top10Filters = new List<LegacyXlsAutoFilterCriteria>();
            foreach (LegacyXlsAutoFilterCriteria columnCriteria in criteria) {
                if (CanProjectAsBlankOrValueList(columnCriteria)) {
                    blankOrValueFilters.Add(columnCriteria);
                } else if (CanProjectAsEqualityList(columnCriteria)) {
                    equalityFilters[columnCriteria.ColumnId] = columnCriteria.Conditions.Select(condition => condition.Value).ToArray();
                } else if (columnCriteria.Kind == LegacyXlsAutoFilterKind.Blanks) {
                    blankFilters.Add(columnCriteria);
                } else if (columnCriteria.Kind == LegacyXlsAutoFilterKind.NonBlanks) {
                    customFilters.Add(columnCriteria);
                } else if (columnCriteria.IsTop10) {
                    top10Filters.Add(columnCriteria);
                } else {
                    customFilters.Add(columnCriteria);
                }
            }

            sheet.AddAutoFilter(range, equalityFilters.Count == 0 ? null : equalityFilters);

            foreach (LegacyXlsAutoFilterCriteria columnCriteria in blankFilters) {
                sheet.ApplyAutoFilterBlankCriteria(range, columnCriteria.ColumnId);
            }

            foreach (LegacyXlsAutoFilterCriteria columnCriteria in blankOrValueFilters) {
                ProjectBlankOrValueList(sheet, range, columnCriteria);
            }

            foreach (LegacyXlsAutoFilterCriteria columnCriteria in customFilters) {
                sheet.ApplyAutoFilterCustomCriteria(
                    range,
                    columnCriteria.ColumnId,
                    columnCriteria.MatchAll,
                    columnCriteria.Conditions.Select(condition => (ToOperator(condition.Operator), ToValue(columnCriteria, condition))).ToArray());
            }

            foreach (LegacyXlsAutoFilterCriteria columnCriteria in top10Filters) {
                sheet.ApplyAutoFilterTop10Criteria(
                    range,
                    columnCriteria.ColumnId,
                    columnCriteria.Top10Value!.Value,
                    columnCriteria.Top10IsTop,
                    columnCriteria.Top10IsPercent);
            }
        }

        private static bool CanProjectAsEqualityList(LegacyXlsAutoFilterCriteria criteria) {
            return criteria.Kind == LegacyXlsAutoFilterKind.Custom
                && criteria.Conditions.Count > 0
                && (criteria.Conditions.Count == 1 || !criteria.MatchAll)
                && criteria.Conditions.All(condition =>
                    condition.ValueKind != LegacyXlsAutoFilterValueKind.Blank
                    && condition.Operator == LegacyXlsAutoFilterOperator.Equal
                    && !condition.HasTextWildcardPattern);
        }

        private static bool CanProjectAsBlankOrValueList(LegacyXlsAutoFilterCriteria criteria) {
            return criteria.Kind == LegacyXlsAutoFilterKind.Custom
                && criteria.Conditions.Count == 2
                && !criteria.MatchAll
                && criteria.Conditions.Count(condition => condition.ValueKind == LegacyXlsAutoFilterValueKind.Blank) == 1
                && criteria.Conditions.Count(condition =>
                    condition.ValueKind != LegacyXlsAutoFilterValueKind.Blank
                    && condition.Operator == LegacyXlsAutoFilterOperator.Equal
                    && !condition.HasTextWildcardPattern) == 1;
        }

        private static void ProjectBlankOrValueList(ExcelSheet sheet, string range, LegacyXlsAutoFilterCriteria criteria) {
            Worksheet? worksheet = sheet.WorksheetPart.Worksheet;
            if (worksheet == null) {
                return;
            }

            AutoFilter? autoFilter = worksheet.GetFirstChild<AutoFilter>();
            if (autoFilter == null || !string.Equals(autoFilter.Reference?.Value, range, StringComparison.OrdinalIgnoreCase)) {
                sheet.AddAutoFilter(range);
                worksheet = sheet.WorksheetPart.Worksheet;
                autoFilter = worksheet?.GetFirstChild<AutoFilter>();
            }

            if (autoFilter == null) {
                return;
            }

            FilterColumn? existingColumn = autoFilter.Elements<FilterColumn>().FirstOrDefault(column => column.ColumnId?.Value == criteria.ColumnId);
            existingColumn?.Remove();

            LegacyXlsAutoFilterCondition valueCondition = criteria.Conditions.Single(condition => condition.ValueKind != LegacyXlsAutoFilterValueKind.Blank);
            var filterColumn = new FilterColumn { ColumnId = criteria.ColumnId };
            filterColumn.Append(new Filters(
                new Filter { Val = valueCondition.Value }) {
                Blank = true
            });
            autoFilter.Append(filterColumn);
        }

        private static FilterOperatorValues ToOperator(LegacyXlsAutoFilterOperator @operator) {
            return @operator switch {
                LegacyXlsAutoFilterOperator.LessThan => FilterOperatorValues.LessThan,
                LegacyXlsAutoFilterOperator.Equal => FilterOperatorValues.Equal,
                LegacyXlsAutoFilterOperator.LessThanOrEqual => FilterOperatorValues.LessThanOrEqual,
                LegacyXlsAutoFilterOperator.GreaterThan => FilterOperatorValues.GreaterThan,
                LegacyXlsAutoFilterOperator.NotEqual => FilterOperatorValues.NotEqual,
                _ => FilterOperatorValues.GreaterThanOrEqual
            };
        }

        private static string ToValue(LegacyXlsAutoFilterCriteria criteria, LegacyXlsAutoFilterCondition condition) {
            return criteria.Kind == LegacyXlsAutoFilterKind.NonBlanks ? " " : condition.Value;
        }
    }
}
