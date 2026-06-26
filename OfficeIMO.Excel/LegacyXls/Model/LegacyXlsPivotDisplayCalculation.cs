namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the display calculation stored in an SXDI PivotTable data item record.
    /// </summary>
    public enum LegacyXlsPivotDisplayCalculation {
        /// <summary>Display the data item value directly.</summary>
        Value = 0,

        /// <summary>Display the difference from a referenced pivot item.</summary>
        DifferenceFrom = 1,

        /// <summary>Display as a percentage of a referenced pivot item.</summary>
        PercentOf = 2,

        /// <summary>Display as a percentage difference from a referenced pivot item.</summary>
        PercentDifferenceFrom = 3,

        /// <summary>Display as a running total for successive pivot items.</summary>
        RunningTotal = 4,

        /// <summary>Display as a percentage of the containing row total.</summary>
        PercentOfRow = 5,

        /// <summary>Display as a percentage of the containing column total.</summary>
        PercentOfColumn = 6,

        /// <summary>Display as a percentage of the grand total.</summary>
        PercentOfGrandTotal = 7,

        /// <summary>Display as index using row, column, and grand totals.</summary>
        Index = 8
    }
}
