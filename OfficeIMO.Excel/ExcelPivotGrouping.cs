using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes date or numeric grouping metadata for a pivot source field.
    /// </summary>
    public sealed class ExcelPivotGrouping {
        private ExcelPivotGrouping(
            string fieldName,
            GroupByValues groupBy,
            DateTime? startDate,
            DateTime? endDate,
            double? startNumber,
            double? endNumber,
            double? interval,
            bool autoStart,
            bool autoEnd,
            IReadOnlyList<GroupByValues>? generatedDateLevels = null) {
            FieldName = string.IsNullOrWhiteSpace(fieldName) ? throw new ArgumentNullException(nameof(fieldName)) : fieldName.Trim();
            GroupBy = groupBy;
            StartDate = startDate;
            EndDate = endDate;
            StartNumber = startNumber;
            EndNumber = endNumber;
            Interval = interval;
            AutoStart = autoStart;
            AutoEnd = autoEnd;
            GeneratedDateLevels = generatedDateLevels ?? Array.Empty<GroupByValues>();
        }

        /// <summary>Creates date grouping metadata for a pivot field.</summary>
        public static ExcelPivotGrouping Date(string fieldName, GroupByValues groupBy, DateTime? startDate = null, DateTime? endDate = null, double? interval = null) {
            if (!IsDateGroupBy(groupBy)) {
                throw new ArgumentException("Date grouping must use days, months, quarters, years, hours, minutes, or seconds.", nameof(groupBy));
            }
            if (interval.HasValue && interval.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(interval), "Grouping interval must be greater than zero.");
            }

            return new ExcelPivotGrouping(fieldName, groupBy, startDate, endDate, null, null, interval, startDate == null, endDate == null);
        }

        /// <summary>Creates generated date hierarchy fields for a pivot field, such as years, quarters, and months.</summary>
        public static ExcelPivotGrouping DateHierarchy(string fieldName, params GroupByValues[] levels) {
            return DateHierarchy(fieldName, levels, null, null);
        }

        /// <summary>Creates generated date hierarchy fields for a pivot field, such as years, quarters, and months.</summary>
        public static ExcelPivotGrouping DateHierarchy(string fieldName, IEnumerable<GroupByValues> levels, DateTime? startDate = null, DateTime? endDate = null) {
            var normalizedLevels = NormalizeDateHierarchyLevels(levels);
            return new ExcelPivotGrouping(
                fieldName,
                normalizedLevels[0],
                startDate,
                endDate,
                null,
                null,
                null,
                startDate == null,
                endDate == null,
                normalizedLevels);
        }

        /// <summary>Creates numeric range grouping metadata for a pivot field.</summary>
        public static ExcelPivotGrouping Number(string fieldName, double interval, double? startNumber = null, double? endNumber = null) {
            if (interval <= 0) {
                throw new ArgumentOutOfRangeException(nameof(interval), "Grouping interval must be greater than zero.");
            }

            return new ExcelPivotGrouping(fieldName, GroupByValues.Range, null, null, startNumber, endNumber, interval, startNumber == null, endNumber == null);
        }

        /// <summary>Gets the source field name.</summary>
        public string FieldName { get; }

        /// <summary>Gets the OpenXML grouping mode.</summary>
        public GroupByValues GroupBy { get; }

        /// <summary>Gets the optional date grouping start.</summary>
        public DateTime? StartDate { get; }

        /// <summary>Gets the optional date grouping end.</summary>
        public DateTime? EndDate { get; }

        /// <summary>Gets the optional numeric grouping start.</summary>
        public double? StartNumber { get; }

        /// <summary>Gets the optional numeric grouping end.</summary>
        public double? EndNumber { get; }

        /// <summary>Gets the grouping interval.</summary>
        public double? Interval { get; }

        /// <summary>Gets whether Excel should infer the start of the grouping range.</summary>
        public bool AutoStart { get; }

        /// <summary>Gets whether Excel should infer the end of the grouping range.</summary>
        public bool AutoEnd { get; }

        /// <summary>Gets generated date hierarchy levels requested for this source field.</summary>
        public IReadOnlyList<GroupByValues> GeneratedDateLevels { get; }

        internal bool IsDateGrouping => IsDateGroupBy(GroupBy);

        internal bool HasGeneratedDateLevels => GeneratedDateLevels.Count > 0;

        private static IReadOnlyList<GroupByValues> NormalizeDateHierarchyLevels(IEnumerable<GroupByValues> levels) {
            if (levels == null) throw new ArgumentNullException(nameof(levels));

            var list = new List<GroupByValues>();
            var used = new HashSet<GroupByValues>();
            foreach (var level in levels) {
                if (!IsDateGroupBy(level)) {
                    throw new ArgumentException("Date hierarchy levels must use days, months, quarters, years, hours, minutes, or seconds.", nameof(levels));
                }

                if (!used.Add(level)) {
                    throw new ArgumentException($"Date hierarchy level '{level}' was specified more than once.", nameof(levels));
                }

                list.Add(level);
            }

            if (list.Count == 0) {
                throw new ArgumentException("At least one date hierarchy level must be specified.", nameof(levels));
            }

            return list;
        }

        private static bool IsDateGroupBy(GroupByValues groupBy) {
            return groupBy == GroupByValues.Seconds
                || groupBy == GroupByValues.Minutes
                || groupBy == GroupByValues.Hours
                || groupBy == GroupByValues.Days
                || groupBy == GroupByValues.Months
                || groupBy == GroupByValues.Quarters
                || groupBy == GroupByValues.Years;
        }
    }
}
