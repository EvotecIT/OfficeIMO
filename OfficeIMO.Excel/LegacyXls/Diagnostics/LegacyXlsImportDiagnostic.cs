namespace OfficeIMO.Excel.LegacyXls.Diagnostics {
    /// <summary>
    /// Describes an import issue, unsupported feature, or compatibility note discovered in a legacy XLS file.
    /// </summary>
    public sealed class LegacyXlsImportDiagnostic {
        /// <summary>
        /// Creates a diagnostic for a legacy XLS import issue or feature note.
        /// </summary>
        /// <param name="severity">Diagnostic severity.</param>
        /// <param name="code">Stable diagnostic code.</param>
        /// <param name="message">Human-readable diagnostic message.</param>
        /// <param name="sheetName">Optional worksheet name associated with the diagnostic.</param>
        /// <param name="recordOffset">Optional byte offset of the related BIFF record.</param>
        /// <param name="recordType">Optional BIFF record type identifier.</param>
        /// <param name="detailCode">Optional stable detail key for grouped import reports.</param>
        public LegacyXlsImportDiagnostic(
            LegacyXlsDiagnosticSeverity severity,
            string code,
            string message,
            string? sheetName = null,
            int? recordOffset = null,
            ushort? recordType = null,
            string? detailCode = null) {
            Severity = severity;
            Code = code;
            Message = message;
            SheetName = sheetName;
            RecordOffset = recordOffset;
            RecordType = recordType;
            DetailCode = detailCode;
        }

        /// <summary>
        /// Gets the diagnostic severity.
        /// </summary>
        public LegacyXlsDiagnosticSeverity Severity { get; }

        /// <summary>
        /// Gets the stable diagnostic code.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Gets the human-readable diagnostic message.
        /// </summary>
        public string Message { get; }

        /// <summary>
        /// Gets the worksheet name associated with the diagnostic, when known.
        /// </summary>
        public string? SheetName { get; }

        /// <summary>
        /// Gets the byte offset of the related BIFF record, when known.
        /// </summary>
        public int? RecordOffset { get; }

        /// <summary>
        /// Gets the BIFF record type identifier, when known.
        /// </summary>
        public ushort? RecordType { get; }

        /// <summary>
        /// Gets a stable detail key for report grouping, when available.
        /// </summary>
        public string? DetailCode { get; }

        /// <summary>
        /// Returns a compact diagnostic string for logs and test output.
        /// </summary>
        public override string ToString() {
            string location = SheetName == null ? string.Empty : $" [{SheetName}]";
            string record = RecordType == null ? string.Empty : $" record=0x{RecordType.Value:X4}";
            string offset = RecordOffset == null ? string.Empty : $" offset={RecordOffset.Value}";
            return $"{Severity}: {Code}{location}{record}{offset}: {Message}";
        }
    }
}
