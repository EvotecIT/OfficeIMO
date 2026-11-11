namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Writes the stylesheet when style edits deferred its save. Call it before a save, clone, snapshot or copy
        /// reads part streams and before the package is handed out; library reads use the loaded stylesheet.
        /// </summary>
        internal void CommitPendingStylesheet() => ExcelSheet.CommitPendingStylesheet(_workBookPartValue?.WorkbookStylesPart);
    }
}
