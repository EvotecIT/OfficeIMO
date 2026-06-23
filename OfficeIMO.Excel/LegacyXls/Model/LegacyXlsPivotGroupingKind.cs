namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the grouping criteria stored in an SXRng PivotTable grouping record.
    /// </summary>
    public enum LegacyXlsPivotGroupingKind {
        /// <summary>Group by numeric value.</summary>
        Numeric = 0,

        /// <summary>Group by seconds.</summary>
        Seconds = 1,

        /// <summary>Group by minutes.</summary>
        Minutes = 2,

        /// <summary>Group by hours.</summary>
        Hours = 3,

        /// <summary>Group by days.</summary>
        Days = 4,

        /// <summary>Group by months.</summary>
        Months = 5,

        /// <summary>Group by quarters.</summary>
        Quarters = 6,

        /// <summary>Group by years.</summary>
        Years = 7
    }
}
