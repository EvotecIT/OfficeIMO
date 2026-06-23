namespace OfficeIMO.Excel {
    /// <summary>
    /// Specifies which part of a pivot table output range should be targeted.
    /// </summary>
    public enum ExcelPivotRangeTarget {
        /// <summary>
        /// Use the full pivot table output range.
        /// </summary>
        WholeTable,

        /// <summary>
        /// Use the data body approximation by excluding the first row and first column when possible.
        /// </summary>
        DataBody
    }
}
