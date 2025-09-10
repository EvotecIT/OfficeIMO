namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls how worksheet names are validated when adding sheets.
    /// </summary>
    public enum SheetNameValidationMode {
        /// <summary>
        /// Do not alter the provided name. The caller is responsible for ensuring it obeys Excel rules.
        /// </summary>
        None = 0,
        /// <summary>
        /// Coerce the provided name into a valid Excel worksheet name by applying the rules:
        /// - Replace invalid characters (: \ / ? * [ ]) with underscore.
        /// - Trim leading/trailing apostrophes and whitespace.
        /// - Truncate to 31 characters.
        /// - Ensure non-empty; fall back to "Sheet".
        /// - Ensure uniqueness across the workbook by appending " (2)", " (3)", etc.
        /// </summary>
        Sanitize = 1,
        /// <summary>
        /// Enforce Excel rules strictly; throw if the name is empty, too long (&gt; 31), contains invalid
        /// characters (: \ / ? * [ ]), or would collide (case-insensitive) with an existing sheet.
        /// </summary>
        Strict = 2
    }
}

