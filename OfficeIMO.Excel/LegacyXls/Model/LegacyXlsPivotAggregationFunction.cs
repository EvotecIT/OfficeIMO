namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the aggregation function stored in an SXDI PivotTable data item record.
    /// </summary>
    public enum LegacyXlsPivotAggregationFunction {
        /// <summary>Sum of values.</summary>
        Sum = 0,

        /// <summary>Count of values.</summary>
        Count = 1,

        /// <summary>Average of values.</summary>
        Average = 2,

        /// <summary>Maximum value.</summary>
        Max = 3,

        /// <summary>Minimum value.</summary>
        Min = 4,

        /// <summary>Product of values.</summary>
        Product = 5,

        /// <summary>Count of numeric values.</summary>
        CountNumbers = 6,

        /// <summary>Sample standard deviation.</summary>
        StandardDeviationSample = 7,

        /// <summary>Population standard deviation.</summary>
        StandardDeviationPopulation = 8,

        /// <summary>Sample variance.</summary>
        VarianceSample = 9,

        /// <summary>Population variance.</summary>
        VariancePopulation = 10
    }
}
