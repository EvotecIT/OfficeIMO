namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls validation behavior for workbook defined names (Named Ranges) and A1 range bounds.
    /// </summary>
    public enum NameValidationMode {
        /// <summary>
        /// Adjust invalid input into a legal Excel defined name (replace invalid characters with underscore,
        /// ensure a valid starting character, trim length). Names that resemble A1/R1C1 cell addresses are prefixed
        /// with an underscore. Out-of-bounds A1 ranges are clamped to Excel limits.
        /// </summary>
        Sanitize = 0,
        /// <summary>
        /// Enforce Excel rules strictly: throws when the name is invalid or the A1 range exceeds Excel bounds.
        /// </summary>
        Strict = 1
    }
}
