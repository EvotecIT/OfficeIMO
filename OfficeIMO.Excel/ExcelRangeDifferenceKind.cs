namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes the kind of difference found while comparing worksheet ranges.
    /// </summary>
    public enum ExcelRangeDifferenceKind {
        /// <summary>A cell exists in both ranges but has a different value.</summary>
        ValueMismatch,

        /// <summary>The compared coordinate exists only in the right range.</summary>
        MissingFromLeft,

        /// <summary>The compared coordinate exists only in the left range.</summary>
        MissingFromRight
    }
}
