namespace OfficeIMO.Excel.Pdf {
    /// <summary>How much detail an Excel PDF conversion report gives about clipped worksheet cells.</summary>
    public enum ExcelPdfDiagnosticsDetail {
        /// <summary>
        /// One diagnostic per contiguous range of clipped cells for each sheet and code, with the range, the cell count
        /// and sample cells. This is the default.
        /// </summary>
        Summary = 0,

        /// <summary>One diagnostic per clipped cell, located by its PDF table cell, page and position.</summary>
        PerCell = 1
    }
}
