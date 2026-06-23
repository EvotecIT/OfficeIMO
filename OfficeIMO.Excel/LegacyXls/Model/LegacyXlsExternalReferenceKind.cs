namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies the supporting-link category represented by a legacy XLS SupBook record.
    /// </summary>
    public enum LegacyXlsExternalReferenceKind {
        /// <summary>
        /// The supporting-link kind could not be classified by this import phase.
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// The supporting link references the current workbook.
        /// </summary>
        Self,

        /// <summary>
        /// The supporting link references the same sheet.
        /// </summary>
        SameSheet,

        /// <summary>
        /// The supporting link references an external workbook.
        /// </summary>
        ExternalWorkbook,

        /// <summary>
        /// The supporting link references an Excel add-in function source.
        /// </summary>
        AddIn,

        /// <summary>
        /// The supporting link references a DDE or OLE data source.
        /// </summary>
        DdeOrOle,

        /// <summary>
        /// The supporting link is an unused placeholder.
        /// </summary>
        Unused
    }
}
