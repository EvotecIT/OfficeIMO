namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for creating a reusable subtotal summary block from worksheet rows.
    /// </summary>
    public sealed class ExcelSubtotalOptions {
        /// <summary>Header row that contains source labels.</summary>
        public int HeaderRow { get; set; } = 1;

        /// <summary>First source data row.</summary>
        public int DataStartRow { get; set; } = 2;

        /// <summary>Last source data row.</summary>
        public int DataEndRow { get; set; }

        /// <summary>Column containing group keys.</summary>
        public int GroupColumn { get; set; }

        /// <summary>Columns to aggregate into subtotal formulas.</summary>
        public IReadOnlyList<int> ValueColumns { get; set; } = Array.Empty<int>();

        /// <summary>First row for the generated summary block. Defaults to two rows below <see cref="DataEndRow"/>.</summary>
        public int? SummaryStartRow { get; set; }

        /// <summary>Function used by generated SUBTOTAL formulas.</summary>
        public ExcelSubtotalFunction Function { get; set; } = ExcelSubtotalFunction.Sum;

        /// <summary>Whether to write a header row for the summary block.</summary>
        public bool IncludeHeader { get; set; } = true;

        /// <summary>Whether to add a grand total row after group subtotal rows.</summary>
        public bool IncludeGrandTotal { get; set; } = true;

        /// <summary>Whether to apply Excel row outline metadata to detail groups.</summary>
        public bool OutlineDetailRows { get; set; } = true;

        /// <summary>Hide detail rows when applying outline metadata.</summary>
        public bool HideDetailRows { get; set; }

        /// <summary>Outline level used for grouped detail rows.</summary>
        public byte OutlineLevel { get; set; } = 1;

        /// <summary>Text appended to each group key in the subtotal label cell.</summary>
        public string LabelSuffix { get; set; } = " Total";

        /// <summary>Label used when the group key cell is blank.</summary>
        public string BlankGroupLabel { get; set; } = "(blank)";

        /// <summary>Label used for the optional grand total row.</summary>
        public string GrandTotalLabel { get; set; } = "Grand Total";
    }
}
