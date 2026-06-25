namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the item type stored in an SXLI PivotTable line item.
    /// </summary>
    public enum LegacyXlsPivotLineItemType {
        /// <summary>Regular data item.</summary>
        Data = 0,

        /// <summary>Default subtotal item.</summary>
        Default = 1,

        /// <summary>Sum subtotal item.</summary>
        Sum = 2,

        /// <summary>Count values subtotal item.</summary>
        CountValues = 3,

        /// <summary>Count numbers subtotal item.</summary>
        CountNumbers = 4,

        /// <summary>Average subtotal item.</summary>
        Average = 5,

        /// <summary>Maximum subtotal item.</summary>
        Max = 6,

        /// <summary>Minimum subtotal item.</summary>
        Min = 7,

        /// <summary>Product subtotal item.</summary>
        Product = 8,

        /// <summary>Standard deviation subtotal item.</summary>
        StandardDeviation = 9,

        /// <summary>Population standard deviation subtotal item.</summary>
        StandardDeviationPopulation = 10,

        /// <summary>Variance subtotal item.</summary>
        Variance = 11,

        /// <summary>Population variance subtotal item.</summary>
        VariancePopulation = 12,

        /// <summary>Grand total item.</summary>
        GrandTotal = 13,

        /// <summary>Blank pivot line.</summary>
        BlankLine = 14
    }
}
