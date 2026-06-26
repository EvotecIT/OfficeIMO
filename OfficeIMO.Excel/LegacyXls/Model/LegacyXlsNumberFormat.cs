namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a number format declared by a legacy XLS FORMAT record.
    /// </summary>
    public sealed class LegacyXlsNumberFormat {
        /// <summary>
        /// Creates a parsed legacy XLS number format.
        /// </summary>
        /// <param name="formatId">Legacy IFmt identifier.</param>
        /// <param name="formatCode">Excel number format code.</param>
        public LegacyXlsNumberFormat(ushort formatId, string formatCode) {
            FormatId = formatId;
            FormatCode = formatCode ?? throw new ArgumentNullException(nameof(formatCode));
        }

        /// <summary>
        /// Gets the legacy IFmt identifier.
        /// </summary>
        public ushort FormatId { get; }

        /// <summary>
        /// Gets the Excel number format code.
        /// </summary>
        public string FormatCode { get; }
    }
}
