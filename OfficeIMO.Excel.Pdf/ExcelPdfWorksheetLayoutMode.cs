namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// Selects how worksheet content is arranged during first-party PDF export.
    /// </summary>
    public enum ExcelPdfWorksheetLayoutMode {
        /// <summary>
        /// Paints cells and drawing-layer objects in worksheet coordinates. This is the
        /// default fidelity mode and keeps table text, links, and form controls native.
        /// </summary>
        WorksheetCanvas = 0,

        /// <summary>
        /// Emits charts, free-standing images, and worksheet tables as document-flow blocks.
        /// This compatibility mode is useful for report-style exports that intentionally
        /// reflow workbook content instead of preserving worksheet geometry.
        /// </summary>
        FlowTable = 1
    }
}
