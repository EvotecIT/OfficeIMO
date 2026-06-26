namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the item type stored in an SXVI PivotTable item record.
    /// </summary>
    public enum LegacyXlsPivotItemType {
        /// <summary>Regular data item.</summary>
        Data = 0,

        /// <summary>Default subtotal item.</summary>
        Default = 1,

        /// <summary>Sum subtotal item.</summary>
        Sum = 2,

        /// <summary>Count values subtotal item.</summary>
        CountValues = 3,

        /// <summary>Average subtotal item.</summary>
        Average = 4,

        /// <summary>Maximum subtotal item.</summary>
        Max = 5,

        /// <summary>Minimum subtotal item.</summary>
        Min = 6,

        /// <summary>Product subtotal item.</summary>
        Product = 7,

        /// <summary>Count numbers subtotal item.</summary>
        CountNumbers = 8,

        /// <summary>Standard deviation subtotal item.</summary>
        StandardDeviation = 9,

        /// <summary>Population standard deviation subtotal item.</summary>
        StandardDeviationPopulation = 10,

        /// <summary>Variance subtotal item.</summary>
        Variance = 11,

        /// <summary>Population variance subtotal item.</summary>
        VariancePopulation = 12
    }
}
