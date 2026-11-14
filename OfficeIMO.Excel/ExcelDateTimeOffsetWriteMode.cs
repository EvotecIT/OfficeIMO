namespace OfficeIMO.Excel {
    /// <summary>
    /// Selects how <see cref="DateTimeOffset"/> values are written to worksheet cells.
    /// </summary>
    public enum ExcelDateTimeOffsetWriteMode {
        /// <summary>
        /// Writes ISO 8601 round-trip text, such as <c>2026-03-14T09:26:53.5897932+05:45</c>, which keeps the offset
        /// and every fractional digit. Excel has no date type with an offset. This is the default.
        /// </summary>
        Text = 0,

        /// <summary>
        /// Writes an Excel date serial with a date number format. The offset is lost when the value is converted to a
        /// <see cref="DateTime"/>; values Excel cannot hold as a serial, such as dates before 1900, are still written as text.
        /// </summary>
        Serial = 1
    }
}
