namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes the value category of a legacy XLS cell.
    /// </summary>
    public enum LegacyXlsCellValueKind {
        /// <summary>
        /// Empty cell placeholder.
        /// </summary>
        Blank,

        /// <summary>
        /// Text value.
        /// </summary>
        Text,

        /// <summary>
        /// Numeric value.
        /// </summary>
        Number,

        /// <summary>
        /// Boolean value.
        /// </summary>
        Boolean,

        /// <summary>
        /// Legacy Excel error value.
        /// </summary>
        Error
    }
}
