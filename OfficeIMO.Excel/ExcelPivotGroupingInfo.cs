using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes grouping metadata read from a pivot cache field.
    /// </summary>
    public sealed class ExcelPivotGroupingInfo {
        /// <summary>Creates pivot grouping readback information.</summary>
        public ExcelPivotGroupingInfo(
            string fieldName,
            GroupByValues? groupBy,
            DateTime? startDate,
            DateTime? endDate,
            double? startNumber,
            double? endNumber,
            double? interval,
            bool? autoStart,
            bool? autoEnd,
            IReadOnlyList<string>? groupItems = null,
            uint? baseFieldIndex = null,
            uint? parentFieldIndex = null) {
            FieldName = fieldName;
            GroupBy = groupBy;
            StartDate = startDate;
            EndDate = endDate;
            StartNumber = startNumber;
            EndNumber = endNumber;
            Interval = interval;
            AutoStart = autoStart;
            AutoEnd = autoEnd;
            GroupItems = groupItems ?? Array.Empty<string>();
            BaseFieldIndex = baseFieldIndex;
            ParentFieldIndex = parentFieldIndex;
        }

        /// <summary>Gets the source field name.</summary>
        public string FieldName { get; }

        /// <summary>Gets the grouping mode.</summary>
        public GroupByValues? GroupBy { get; }

        /// <summary>Gets the date grouping start, if present.</summary>
        public DateTime? StartDate { get; }

        /// <summary>Gets the date grouping end, if present.</summary>
        public DateTime? EndDate { get; }

        /// <summary>Gets the numeric grouping start, if present.</summary>
        public double? StartNumber { get; }

        /// <summary>Gets the numeric grouping end, if present.</summary>
        public double? EndNumber { get; }

        /// <summary>Gets the grouping interval, if present.</summary>
        public double? Interval { get; }

        /// <summary>Gets whether Excel should infer the grouping start.</summary>
        public bool? AutoStart { get; }

        /// <summary>Gets whether Excel should infer the grouping end.</summary>
        public bool? AutoEnd { get; }

        /// <summary>Gets explicit grouping item labels or values, if present.</summary>
        public IReadOnlyList<string> GroupItems { get; }

        /// <summary>Gets the source field index used as the base for this grouping, if present.</summary>
        public uint? BaseFieldIndex { get; }

        /// <summary>Gets the parent grouping field index, if present.</summary>
        public uint? ParentFieldIndex { get; }
    }
}
