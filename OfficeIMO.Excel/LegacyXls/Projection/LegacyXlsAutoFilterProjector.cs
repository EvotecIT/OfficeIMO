using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static class LegacyXlsAutoFilterProjector {
        internal static void Project(ExcelSheet sheet, string range, IReadOnlyList<LegacyXlsAutoFilterCriteria> criteria) {
            if (criteria.Count == 0 || string.IsNullOrWhiteSpace(range)) {
                return;
            }

            var equalityFilters = new Dictionary<uint, IEnumerable<string>>();
            var customFilters = new List<LegacyXlsAutoFilterCriteria>();
            foreach (LegacyXlsAutoFilterCriteria columnCriteria in criteria) {
                if (CanProjectAsEqualityList(columnCriteria)) {
                    equalityFilters[columnCriteria.ColumnId] = columnCriteria.Conditions.Select(condition => condition.Value).ToArray();
                } else {
                    customFilters.Add(columnCriteria);
                }
            }

            sheet.AddAutoFilter(range, equalityFilters.Count == 0 ? null : equalityFilters);

            foreach (LegacyXlsAutoFilterCriteria columnCriteria in customFilters) {
                sheet.ApplyAutoFilterCustomCriteria(
                    range,
                    columnCriteria.ColumnId,
                    columnCriteria.MatchAll,
                    columnCriteria.Conditions.Select(condition => (ToOperator(condition.Operator), condition.Value)).ToArray());
            }
        }

        private static bool CanProjectAsEqualityList(LegacyXlsAutoFilterCriteria criteria) {
            return criteria.Conditions.Count > 0
                && !criteria.MatchAll
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
    }
}
