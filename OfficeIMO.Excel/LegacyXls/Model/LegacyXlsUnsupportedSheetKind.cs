namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Identifies legacy XLS sheet entries that are discovered but not imported as worksheets.
    /// </summary>
    public enum LegacyXlsUnsupportedSheetKind {
        /// <summary>
        /// The sheet type is not recognized by this import phase.
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// The sheet is an Excel macro sheet.
        /// </summary>
        MacroSheet,

        /// <summary>
        /// The sheet is a chart sheet.
        /// </summary>
        ChartSheet,

        /// <summary>
        /// The sheet is a VBA module sheet.
        /// </summary>
        VbaModuleSheet,

        /// <summary>
        /// The sheet is a dialog sheet identified by the WsBool fDialog flag.
        /// </summary>
        DialogSheet
    }
}
