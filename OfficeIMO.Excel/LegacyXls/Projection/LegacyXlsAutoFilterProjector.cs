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
            var customFilters = new List<LegacyXlsAutoFilterCriteria>();
            var top10Filters = new List<LegacyXlsAutoFilterCriteria>();
            foreach (LegacyXlsAutoFilterCriteria columnCriteria in criteria) {
                if (CanProjectAsEqualityList(columnCriteria)) {
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
                && criteria.Conditions.All(condition => condition.Operator == LegacyXlsAutoFilterOperator.Equal);
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
