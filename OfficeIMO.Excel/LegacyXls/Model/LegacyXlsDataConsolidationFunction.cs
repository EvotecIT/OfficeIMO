namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the aggregation function carried by a BIFF DCon data-consolidation settings record.
    /// </summary>
    public enum LegacyXlsDataConsolidationFunction {
        /// <summary>Average source values.</summary>
        Average = 0,

        /// <summary>Count numeric source values.</summary>
        CountNumbers = 1,

        /// <summary>Count source values.</summary>
        Count = 2,

        /// <summary>Use the maximum source value.</summary>
        Maximum = 3,

        /// <summary>Use the minimum source value.</summary>
        Minimum = 4,

        /// <summary>Multiply source values.</summary>
        Product = 5,

        /// <summary>Calculate sample standard deviation.</summary>
        StandardDeviation = 6,

        /// <summary>Calculate population standard deviation.</summary>
        StandardDeviationP = 7,

        /// <summary>Sum source values.</summary>
        Sum = 8,

        /// <summary>Calculate sample variance.</summary>
        Variance = 9,

        /// <summary>Calculate population variance.</summary>
        VarianceP = 10
    }
}
