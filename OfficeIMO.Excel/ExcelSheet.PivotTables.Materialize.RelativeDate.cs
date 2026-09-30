using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly PivotFilterValues[] MaterializedPivotRelativeDateTypes = {
            PivotFilterValues.Yesterday, PivotFilterValues.Today, PivotFilterValues.Tomorrow,
            PivotFilterValues.LastWeek, PivotFilterValues.ThisWeek, PivotFilterValues.NextWeek,
            PivotFilterValues.LastMonth, PivotFilterValues.ThisMonth, PivotFilterValues.NextMonth,
            PivotFilterValues.LastQuarter, PivotFilterValues.ThisQuarter, PivotFilterValues.NextQuarter,
            PivotFilterValues.LastYear, PivotFilterValues.ThisYear, PivotFilterValues.NextYear,
            PivotFilterValues.YearToDate
        };

        private static bool IsMaterializedPivotRelativeDateFilter(PivotFilter filter)
            => filter.Type != null && Array.IndexOf(MaterializedPivotRelativeDateTypes, filter.Type.Value) >= 0;

        private static DynamicFilter QualifiedMaterializedPivotRelativeDateFilter(PivotFilter filter) {
            if (filter.MeasureField != null || filter.StringValue1 != null || filter.StringValue2 != null)
                throw new NotSupportedException("The relative-date filter has no qualified saved predicate.");
            var columns = filter.AutoFilter?.Elements<FilterColumn>().ToArray() ?? Array.Empty<FilterColumn>();
            if (filter.AutoFilter?.ChildElements.Count != 1 || columns.Length != 1
                || columns[0].ColumnId?.Value != 0 || columns[0].ChildElements.Count != 1
                || columns[0].GetFirstChild<DynamicFilter>() is not DynamicFilter dynamic
                || dynamic.ChildElements.Count != 0 || dynamic.Type?.Value != ResolveDynamicFilterType(filter.Type!.Value)
                || dynamic.GetAttributes().Count is not (1 or 3)
                || (dynamic.Val == null) != (dynamic.MaxVal == null)
                || (dynamic.Val != null && (double.IsNaN(dynamic.Val.Value) || double.IsInfinity(dynamic.Val.Value)
                    || double.IsNaN(dynamic.MaxVal!.Value) || double.IsInfinity(dynamic.MaxVal.Value))))
                throw new NotSupportedException("The relative-date filter has no qualified saved dynamic predicate.");
            return dynamic;
        }

        private static bool TryMaterializedPivotRelativeDateBounds(PivotFilterValues type, DateTime today,
            out DateTime start, out DateTime end) {
            DateTime week = today.AddDays(-(int)today.DayOfWeek);
            DateTime month = new(today.Year, today.Month, 1);
            DateTime quarter = new(today.Year, ((today.Month - 1) / 3) * 3 + 1, 1);
            DateTime year = new(today.Year, 1, 1);
            if (type == PivotFilterValues.Yesterday) { start = today.AddDays(-1); end = today; }
            else if (type == PivotFilterValues.Today) { start = today; end = today.AddDays(1); }
            else if (type == PivotFilterValues.Tomorrow) { start = today.AddDays(1); end = today.AddDays(2); }
            else if (type == PivotFilterValues.LastWeek) { start = week.AddDays(-7); end = week; }
            else if (type == PivotFilterValues.ThisWeek) { start = week; end = week.AddDays(7); }
            else if (type == PivotFilterValues.NextWeek) { start = week.AddDays(7); end = week.AddDays(14); }
            else if (type == PivotFilterValues.LastMonth) { start = month.AddMonths(-1); end = month; }
            else if (type == PivotFilterValues.ThisMonth) { start = month; end = month.AddMonths(1); }
            else if (type == PivotFilterValues.NextMonth) { start = month.AddMonths(1); end = month.AddMonths(2); }
            else if (type == PivotFilterValues.LastQuarter) { start = quarter.AddMonths(-3); end = quarter; }
            else if (type == PivotFilterValues.ThisQuarter) { start = quarter; end = quarter.AddMonths(3); }
            else if (type == PivotFilterValues.NextQuarter) { start = quarter.AddMonths(3); end = quarter.AddMonths(6); }
            else if (type == PivotFilterValues.LastYear) { start = year.AddYears(-1); end = year; }
            else if (type == PivotFilterValues.ThisYear) { start = year; end = year.AddYears(1); }
            else if (type == PivotFilterValues.NextYear) { start = year.AddYears(1); end = year.AddYears(2); }
            else if (type == PivotFilterValues.YearToDate) { start = year; end = today.AddDays(1); }
            else { start = end = default; return false; }
            return true;
        }

        private static void UpdateMaterializedPivotRelativeDateBounds(PivotTableDefinition definition,
            DateTime referenceDate, ExcelDateSystem dateSystem) {
            foreach (var filter in definition.PivotFilters?.Elements<PivotFilter>()
                .Where(IsMaterializedPivotRelativeDateFilter) ?? Enumerable.Empty<PivotFilter>()) {
                if (!TryMaterializedPivotRelativeDateBounds(filter.Type!.Value, referenceDate,
                    out DateTime start, out DateTime end)) continue;
                DynamicFilter dynamic = QualifiedMaterializedPivotRelativeDateFilter(filter);
                dynamic.Val = ExcelPivotCacheDateCodec.ToSerial(start, dateSystem);
                dynamic.MaxVal = ExcelPivotCacheDateCodec.ToSerial(end, dateSystem);
            }
        }
    }
}
