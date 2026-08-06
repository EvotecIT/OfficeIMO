namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls how AddTable validates the provided table name.
    /// </summary>
    public enum ExcelTableNameValidationMode {
        /// <summary>
        /// Replace invalid characters, normalize, and ensure uniqueness.
        /// </summary>
        Sanitize = 0,
        /// <summary>
        /// Throw descriptive exceptions when the name doesn't meet requirements.
        /// </summary>
        Strict = 1
    }
}

