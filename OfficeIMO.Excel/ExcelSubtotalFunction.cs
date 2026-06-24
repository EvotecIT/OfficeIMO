namespace OfficeIMO.Excel {
    /// <summary>
    /// Aggregation function used by worksheet subtotal summary rows.
    /// </summary>
    public enum ExcelSubtotalFunction {
        /// <summary>Average visible numeric values.</summary>
        Average,
        /// <summary>Count visible numeric values.</summary>
        Count,
        /// <summary>Count visible non-blank values.</summary>
        CountNonBlank,
        /// <summary>Maximum visible numeric value.</summary>
        Max,
        /// <summary>Minimum visible numeric value.</summary>
        Min,
        /// <summary>Sum visible numeric values.</summary>
        Sum
    }
}
