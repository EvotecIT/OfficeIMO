namespace OfficeIMO.Excel.LegacyXls.Diagnostics {
    /// <summary>
    /// Severity for diagnostics produced while importing legacy XLS workbooks.
    /// </summary>
    public enum LegacyXlsDiagnosticSeverity {
        /// <summary>
        /// Informational note about a feature or record that did not block import.
        /// </summary>
        Info,

        /// <summary>
        /// Recoverable issue that may affect fidelity but still allows import to continue.
        /// </summary>
        Warning,

        /// <summary>
        /// Non-recoverable issue that prevents the requested legacy workbook content from being imported.
        /// </summary>
        Error
    }
}
