namespace OfficeIMO.Excel {
    /// <summary>
    /// Options controlling worksheet/range value comparisons.
    /// </summary>
    public sealed class ExcelRangeCompareOptions {
        /// <summary>
        /// Treats null values and empty strings as equal. Enabled by default to match common reporting expectations.
        /// </summary>
        public bool TreatNullAndEmptyStringAsEqual { get; set; } = true;

        /// <summary>
        /// Trims string values before comparing.
        /// </summary>
        public bool TrimStrings { get; set; }

        /// <summary>
        /// Compares string values using ordinal case-insensitive comparison.
        /// </summary>
        public bool IgnoreCase { get; set; }
    }
}
