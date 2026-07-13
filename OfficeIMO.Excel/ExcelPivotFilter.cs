using System;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a pivot table label or value filter to author.
    /// </summary>
    public sealed class ExcelPivotFilter {
        private ExcelPivotFilter(
            string fieldName,
            PivotFilterValues type,
            string? value1,
            string? value2,
            string? dataFieldName,
            string? name,
            string? description,
            bool? isTop = null,
            bool? isPercent = null,
            string? filterValue = null,
            DateTime? dateValue1 = null,
            DateTime? dateValue2 = null) {
            FieldName = string.IsNullOrWhiteSpace(fieldName) ? throw new ArgumentNullException(nameof(fieldName)) : fieldName.Trim();
            Type = type;
            Value1 = string.IsNullOrWhiteSpace(value1) ? null : value1;
            Value2 = string.IsNullOrWhiteSpace(value2) ? null : value2;
            DataFieldName = string.IsNullOrWhiteSpace(dataFieldName) ? null : dataFieldName!.Trim();
            Name = string.IsNullOrWhiteSpace(name) ? null : name!.Trim();
            Description = string.IsNullOrWhiteSpace(description) ? null : description!.Trim();
            IsTop = isTop;
            IsPercent = isPercent;
            FilterValue = string.IsNullOrWhiteSpace(filterValue) ? null : filterValue;
            DateValue1 = dateValue1;
            DateValue2 = dateValue2;
        }

        /// <summary>Gets the source field name to filter.</summary>
        public string FieldName { get; }

        /// <summary>Gets the Open XML pivot filter type.</summary>
        public PivotFilterValues Type { get; }

        /// <summary>Gets the first filter value.</summary>
        public string? Value1 { get; }

        /// <summary>Gets the second filter value for between-style filters.</summary>
        public string? Value2 { get; }

        /// <summary>Gets the data field used by value filters.</summary>
        public string? DataFieldName { get; }

        /// <summary>Gets the optional filter display name.</summary>
        public string? Name { get; }

        /// <summary>Gets the optional filter description.</summary>
        public string? Description { get; }

        /// <summary>Gets whether a top/bottom filter keeps top values. False means bottom values.</summary>
        public bool? IsTop { get; }

        /// <summary>Gets whether a top/bottom filter uses a percentage threshold.</summary>
        public bool? IsPercent { get; }

        /// <summary>Gets the optional calculated top/bottom filter value threshold.</summary>
        public string? FilterValue { get; }

        internal DateTime? DateValue1 { get; }

        internal DateTime? DateValue2 { get; }

        /// <summary>Creates a label-equals pivot filter.</summary>
        public static ExcelPivotFilter LabelEquals(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionEqual, value, null, name, description);

        /// <summary>Creates a label-not-equals pivot filter.</summary>
        public static ExcelPivotFilter LabelNotEquals(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionNotEqual, value, null, name, description);

        /// <summary>Creates a label-contains pivot filter.</summary>
        public static ExcelPivotFilter LabelContains(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionContains, value, null, name, description);

        /// <summary>Creates a label-not-contains pivot filter.</summary>
        public static ExcelPivotFilter LabelNotContains(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionNotContains, value, null, name, description);

        /// <summary>Creates a label-begins-with pivot filter.</summary>
        public static ExcelPivotFilter LabelBeginsWith(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionBeginsWith, value, null, name, description);

        /// <summary>Creates a label-not-begins-with pivot filter.</summary>
        public static ExcelPivotFilter LabelNotBeginsWith(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionNotBeginsWith, value, null, name, description);

        /// <summary>Creates a label-ends-with pivot filter.</summary>
        public static ExcelPivotFilter LabelEndsWith(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionEndsWith, value, null, name, description);

        /// <summary>Creates a label-not-ends-with pivot filter.</summary>
        public static ExcelPivotFilter LabelNotEndsWith(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionNotEndsWith, value, null, name, description);

        /// <summary>Creates a label-greater-than pivot filter.</summary>
        public static ExcelPivotFilter LabelGreaterThan(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionGreaterThan, value, null, name, description);

        /// <summary>Creates a label-greater-than-or-equal pivot filter.</summary>
        public static ExcelPivotFilter LabelGreaterThanOrEqual(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionGreaterThanOrEqual, value, null, name, description);

        /// <summary>Creates a label-less-than pivot filter.</summary>
        public static ExcelPivotFilter LabelLessThan(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionLessThan, value, null, name, description);

        /// <summary>Creates a label-less-than-or-equal pivot filter.</summary>
        public static ExcelPivotFilter LabelLessThanOrEqual(string fieldName, string value, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionLessThanOrEqual, value, null, name, description);

        /// <summary>Creates a label-between pivot filter.</summary>
        public static ExcelPivotFilter LabelBetween(string fieldName, string from, string to, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionBetween, from, to, name, description);

        /// <summary>Creates a label-not-between pivot filter.</summary>
        public static ExcelPivotFilter LabelNotBetween(string fieldName, string from, string to, string? name = null, string? description = null)
            => Label(fieldName, PivotFilterValues.CaptionNotBetween, from, to, name, description);

        /// <summary>Creates a value-equals pivot filter.</summary>
        public static ExcelPivotFilter ValueEquals(string fieldName, string dataFieldName, double value, string? name = null, string? description = null)
            => Value(fieldName, dataFieldName, PivotFilterValues.ValueEqual, value, null, name, description);

        /// <summary>Creates a value-not-equals pivot filter.</summary>
        public static ExcelPivotFilter ValueNotEquals(string fieldName, string dataFieldName, double value, string? name = null, string? description = null)
            => Value(fieldName, dataFieldName, PivotFilterValues.ValueNotEqual, value, null, name, description);

        /// <summary>Creates a value-greater-than pivot filter.</summary>
        public static ExcelPivotFilter ValueGreaterThan(string fieldName, string dataFieldName, double value, string? name = null, string? description = null)
            => Value(fieldName, dataFieldName, PivotFilterValues.ValueGreaterThan, value, null, name, description);

        /// <summary>Creates a value-greater-than-or-equal pivot filter.</summary>
        public static ExcelPivotFilter ValueGreaterThanOrEqual(string fieldName, string dataFieldName, double value, string? name = null, string? description = null)
            => Value(fieldName, dataFieldName, PivotFilterValues.ValueGreaterThanOrEqual, value, null, name, description);

        /// <summary>Creates a value-less-than pivot filter.</summary>
        public static ExcelPivotFilter ValueLessThan(string fieldName, string dataFieldName, double value, string? name = null, string? description = null)
            => Value(fieldName, dataFieldName, PivotFilterValues.ValueLessThan, value, null, name, description);

        /// <summary>Creates a value-less-than-or-equal pivot filter.</summary>
        public static ExcelPivotFilter ValueLessThanOrEqual(string fieldName, string dataFieldName, double value, string? name = null, string? description = null)
            => Value(fieldName, dataFieldName, PivotFilterValues.ValueLessThanOrEqual, value, null, name, description);

        /// <summary>Creates a value-between pivot filter.</summary>
        public static ExcelPivotFilter ValueBetween(string fieldName, string dataFieldName, double from, double to, string? name = null, string? description = null)
            => Value(fieldName, dataFieldName, PivotFilterValues.ValueBetween, from, to, name, description);

        /// <summary>Creates a value-not-between pivot filter.</summary>
        public static ExcelPivotFilter ValueNotBetween(string fieldName, string dataFieldName, double from, double to, string? name = null, string? description = null)
            => Value(fieldName, dataFieldName, PivotFilterValues.ValueNotBetween, from, to, name, description);

        /// <summary>Creates a top-count pivot filter for the specified data field.</summary>
        public static ExcelPivotFilter TopCount(string fieldName, string dataFieldName, int count, string? name = null, string? description = null)
            => TopBottom(fieldName, dataFieldName, count, isTop: true, isPercent: false, name, description);

        /// <summary>Creates a bottom-count pivot filter for the specified data field.</summary>
        public static ExcelPivotFilter BottomCount(string fieldName, string dataFieldName, int count, string? name = null, string? description = null)
            => TopBottom(fieldName, dataFieldName, count, isTop: false, isPercent: false, name, description);

        /// <summary>Creates a top-percent pivot filter for the specified data field.</summary>
        public static ExcelPivotFilter TopPercent(string fieldName, string dataFieldName, int percent, string? name = null, string? description = null)
            => TopBottom(fieldName, dataFieldName, percent, isTop: true, isPercent: true, name, description);

        /// <summary>Creates a bottom-percent pivot filter for the specified data field.</summary>
        public static ExcelPivotFilter BottomPercent(string fieldName, string dataFieldName, int percent, string? name = null, string? description = null)
            => TopBottom(fieldName, dataFieldName, percent, isTop: false, isPercent: true, name, description);

        /// <summary>Creates a top-sum pivot filter for the specified data field.</summary>
        public static ExcelPivotFilter TopSum(string fieldName, string dataFieldName, double value, string? name = null, string? description = null)
            => TopBottomSum(fieldName, dataFieldName, value, isTop: true, name, description);

        /// <summary>Creates a bottom-sum pivot filter for the specified data field.</summary>
        public static ExcelPivotFilter BottomSum(string fieldName, string dataFieldName, double value, string? name = null, string? description = null)
            => TopBottomSum(fieldName, dataFieldName, value, isTop: false, name, description);

        /// <summary>Creates a date-equals pivot filter.</summary>
        public static ExcelPivotFilter DateEquals(string fieldName, DateTime value, string? name = null, string? description = null)
            => Date(fieldName, PivotFilterValues.DateEqual, value, null, name, description);

        /// <summary>Creates a date-not-equals pivot filter.</summary>
        public static ExcelPivotFilter DateNotEquals(string fieldName, DateTime value, string? name = null, string? description = null)
            => Date(fieldName, PivotFilterValues.DateNotEqual, value, null, name, description);

        /// <summary>Creates a date-newer-than pivot filter.</summary>
        public static ExcelPivotFilter DateNewerThan(string fieldName, DateTime value, string? name = null, string? description = null)
            => Date(fieldName, PivotFilterValues.DateNewerThan, value, null, name, description);

        /// <summary>Creates a date-newer-than-or-equal pivot filter.</summary>
        public static ExcelPivotFilter DateNewerThanOrEqual(string fieldName, DateTime value, string? name = null, string? description = null)
            => Date(fieldName, PivotFilterValues.DateNewerThanOrEqual, value, null, name, description);

        /// <summary>Creates a date-older-than pivot filter.</summary>
        public static ExcelPivotFilter DateOlderThan(string fieldName, DateTime value, string? name = null, string? description = null)
            => Date(fieldName, PivotFilterValues.DateOlderThan, value, null, name, description);

        /// <summary>Creates a date-older-than-or-equal pivot filter.</summary>
        public static ExcelPivotFilter DateOlderThanOrEqual(string fieldName, DateTime value, string? name = null, string? description = null)
            => Date(fieldName, PivotFilterValues.DateOlderThanOrEqual, value, null, name, description);

        /// <summary>Creates a date-between pivot filter.</summary>
        public static ExcelPivotFilter DateBetween(string fieldName, DateTime from, DateTime to, string? name = null, string? description = null)
            => Date(fieldName, PivotFilterValues.DateBetween, from, to, name, description);

        /// <summary>Creates a date-not-between pivot filter.</summary>
        public static ExcelPivotFilter DateNotBetween(string fieldName, DateTime from, DateTime to, string? name = null, string? description = null)
            => Date(fieldName, PivotFilterValues.DateNotBetween, from, to, name, description);

        /// <summary>Creates a dynamic date filter for today.</summary>
        public static ExcelPivotFilter DateToday(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.Today, name, description);

        /// <summary>Creates a dynamic date filter for yesterday.</summary>
        public static ExcelPivotFilter DateYesterday(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.Yesterday, name, description);

        /// <summary>Creates a dynamic date filter for tomorrow.</summary>
        public static ExcelPivotFilter DateTomorrow(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.Tomorrow, name, description);

        /// <summary>Creates a dynamic date filter for this week.</summary>
        public static ExcelPivotFilter DateThisWeek(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.ThisWeek, name, description);

        /// <summary>Creates a dynamic date filter for last week.</summary>
        public static ExcelPivotFilter DateLastWeek(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.LastWeek, name, description);

        /// <summary>Creates a dynamic date filter for next week.</summary>
        public static ExcelPivotFilter DateNextWeek(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.NextWeek, name, description);

        /// <summary>Creates a dynamic date filter for this month.</summary>
        public static ExcelPivotFilter DateThisMonth(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.ThisMonth, name, description);

        /// <summary>Creates a dynamic date filter for last month.</summary>
        public static ExcelPivotFilter DateLastMonth(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.LastMonth, name, description);

        /// <summary>Creates a dynamic date filter for next month.</summary>
        public static ExcelPivotFilter DateNextMonth(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.NextMonth, name, description);

        /// <summary>Creates a dynamic date filter for this quarter.</summary>
        public static ExcelPivotFilter DateThisQuarter(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.ThisQuarter, name, description);

        /// <summary>Creates a dynamic date filter for last quarter.</summary>
        public static ExcelPivotFilter DateLastQuarter(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.LastQuarter, name, description);

        /// <summary>Creates a dynamic date filter for next quarter.</summary>
        public static ExcelPivotFilter DateNextQuarter(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.NextQuarter, name, description);

        /// <summary>Creates a dynamic date filter for this year.</summary>
        public static ExcelPivotFilter DateThisYear(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.ThisYear, name, description);

        /// <summary>Creates a dynamic date filter for last year.</summary>
        public static ExcelPivotFilter DateLastYear(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.LastYear, name, description);

        /// <summary>Creates a dynamic date filter for next year.</summary>
        public static ExcelPivotFilter DateNextYear(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.NextYear, name, description);

        /// <summary>Creates a dynamic date filter for year-to-date.</summary>
        public static ExcelPivotFilter DateYearToDate(string fieldName, string? name = null, string? description = null)
            => DynamicDate(fieldName, PivotFilterValues.YearToDate, name, description);

        /// <summary>Creates a dynamic date filter for a calendar month.</summary>
        public static ExcelPivotFilter DateMonth(string fieldName, int month, string? name = null, string? description = null) {
            switch (month) {
                case 1: return DynamicDate(fieldName, PivotFilterValues.January, name, description);
                case 2: return DynamicDate(fieldName, PivotFilterValues.February, name, description);
                case 3: return DynamicDate(fieldName, PivotFilterValues.March, name, description);
                case 4: return DynamicDate(fieldName, PivotFilterValues.April, name, description);
                case 5: return DynamicDate(fieldName, PivotFilterValues.May, name, description);
                case 6: return DynamicDate(fieldName, PivotFilterValues.June, name, description);
                case 7: return DynamicDate(fieldName, PivotFilterValues.July, name, description);
                case 8: return DynamicDate(fieldName, PivotFilterValues.August, name, description);
                case 9: return DynamicDate(fieldName, PivotFilterValues.September, name, description);
                case 10: return DynamicDate(fieldName, PivotFilterValues.October, name, description);
                case 11: return DynamicDate(fieldName, PivotFilterValues.November, name, description);
                case 12: return DynamicDate(fieldName, PivotFilterValues.December, name, description);
                default: throw new ArgumentOutOfRangeException(nameof(month), "Month must be between 1 and 12.");
            }
        }

        /// <summary>Creates a dynamic date filter for a calendar quarter.</summary>
        public static ExcelPivotFilter DateQuarter(string fieldName, int quarter, string? name = null, string? description = null) {
            switch (quarter) {
                case 1: return DynamicDate(fieldName, PivotFilterValues.Quarter1, name, description);
                case 2: return DynamicDate(fieldName, PivotFilterValues.Quarter2, name, description);
                case 3: return DynamicDate(fieldName, PivotFilterValues.Quarter3, name, description);
                case 4: return DynamicDate(fieldName, PivotFilterValues.Quarter4, name, description);
                default: throw new ArgumentOutOfRangeException(nameof(quarter), "Quarter must be between 1 and 4.");
            }
        }

        /// <summary>Creates a label pivot filter using a specific Open XML pivot filter type.</summary>
        public static ExcelPivotFilter Label(string fieldName, PivotFilterValues type, string value1, string? value2 = null, string? name = null, string? description = null) {
            return new ExcelPivotFilter(fieldName, type, value1, value2, null, name, description);
        }

        /// <summary>Creates a value pivot filter using a specific Open XML pivot filter type.</summary>
        public static ExcelPivotFilter Value(string fieldName, string dataFieldName, PivotFilterValues type, double value1, double? value2 = null, string? name = null, string? description = null) {
            string first = InvariantNumberText.Get(value1);
            string? second = value2.HasValue ? InvariantNumberText.Get(value2.Value) : null;
            return new ExcelPivotFilter(fieldName, type, first, second, dataFieldName, name, description);
        }

        /// <summary>Creates a fixed date pivot filter using a supported Open XML pivot filter type.</summary>
        public static ExcelPivotFilter Date(string fieldName, PivotFilterValues type, DateTime value1, DateTime? value2 = null, string? name = null, string? description = null) {
            if (!IsFixedDateFilter(type)) {
                throw new ArgumentException($"Pivot filter type '{type}' is not a supported fixed date filter.", nameof(type));
            }

            string first = FormatDateFilterValue(value1);
            string? second = value2.HasValue ? FormatDateFilterValue(value2.Value) : null;
            return new ExcelPivotFilter(fieldName, type, first, second, null, name, description, dateValue1: value1, dateValue2: value2);
        }

        /// <summary>Creates a dynamic date pivot filter using a supported Open XML pivot filter type.</summary>
        public static ExcelPivotFilter DynamicDate(string fieldName, PivotFilterValues type, string? name = null, string? description = null) {
            if (!IsDynamicDateFilter(type)) {
                throw new ArgumentException($"Pivot filter type '{type}' is not a supported dynamic date filter.", nameof(type));
            }

            return new ExcelPivotFilter(fieldName, type, null, null, null, name, description);
        }

        private static string FormatDateFilterValue(DateTime value) {
            return InvariantNumberText.Get(value.ToOADate());
        }

        private static bool IsFixedDateFilter(PivotFilterValues type) {
            return type == PivotFilterValues.DateEqual
                || type == PivotFilterValues.DateNotEqual
                || type == PivotFilterValues.DateNewerThan
                || type == PivotFilterValues.DateNewerThanOrEqual
                || type == PivotFilterValues.DateOlderThan
                || type == PivotFilterValues.DateOlderThanOrEqual
                || type == PivotFilterValues.DateBetween
                || type == PivotFilterValues.DateNotBetween;
        }

        private static ExcelPivotFilter TopBottom(string fieldName, string dataFieldName, int value, bool isTop, bool isPercent, string? name, string? description) {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value), "Top/bottom filter values must be greater than zero.");
            if (isPercent && value > 100) throw new ArgumentOutOfRangeException(nameof(value), "Percent filters must be between 1 and 100.");

            var type = isPercent ? PivotFilterValues.Percent : PivotFilterValues.Count;
            string text = value.ToString(CultureInfo.InvariantCulture);
            return new ExcelPivotFilter(fieldName, type, text, null, dataFieldName, name, description, isTop, isPercent);
        }

        private static ExcelPivotFilter TopBottomSum(string fieldName, string dataFieldName, double value, bool isTop, string? name, string? description) {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value), "Top/bottom sum filter values must be greater than zero.");

            string text = InvariantNumberText.Get(value);
            return new ExcelPivotFilter(fieldName, PivotFilterValues.Sum, text, null, dataFieldName, name, description, isTop, false);
        }

        private static bool IsDynamicDateFilter(PivotFilterValues type) {
            return type == PivotFilterValues.Today
                || type == PivotFilterValues.Yesterday
                || type == PivotFilterValues.Tomorrow
                || type == PivotFilterValues.ThisWeek
                || type == PivotFilterValues.LastWeek
                || type == PivotFilterValues.NextWeek
                || type == PivotFilterValues.ThisMonth
                || type == PivotFilterValues.LastMonth
                || type == PivotFilterValues.NextMonth
                || type == PivotFilterValues.ThisQuarter
                || type == PivotFilterValues.LastQuarter
                || type == PivotFilterValues.NextQuarter
                || type == PivotFilterValues.ThisYear
                || type == PivotFilterValues.LastYear
                || type == PivotFilterValues.NextYear
                || type == PivotFilterValues.YearToDate
                || type == PivotFilterValues.January
                || type == PivotFilterValues.February
                || type == PivotFilterValues.March
                || type == PivotFilterValues.April
                || type == PivotFilterValues.May
                || type == PivotFilterValues.June
                || type == PivotFilterValues.July
                || type == PivotFilterValues.August
                || type == PivotFilterValues.September
                || type == PivotFilterValues.October
                || type == PivotFilterValues.November
                || type == PivotFilterValues.December
                || type == PivotFilterValues.Quarter1
                || type == PivotFilterValues.Quarter2
                || type == PivotFilterValues.Quarter3
                || type == PivotFilterValues.Quarter4;
        }
    }
}
