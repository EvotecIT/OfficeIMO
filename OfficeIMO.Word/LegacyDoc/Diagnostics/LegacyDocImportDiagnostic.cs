namespace OfficeIMO.Word.LegacyDoc.Diagnostics {
    /// <summary>
    /// Diagnostic produced while importing a legacy binary Word document.
    /// </summary>
    public sealed class LegacyDocImportDiagnostic {
        /// <summary>
        /// Initializes a new legacy DOC import diagnostic.
        /// </summary>
        /// <param name="code">Stable diagnostic code.</param>
        /// <param name="severity">Diagnostic severity.</param>
        /// <param name="message">Human-readable diagnostic message.</param>
        public LegacyDocImportDiagnostic(string code, LegacyDocDiagnosticSeverity severity, string message) {
            Code = string.IsNullOrWhiteSpace(code) ? throw new ArgumentException("Diagnostic code is required.", nameof(code)) : code;
            Severity = severity;
            Message = message ?? throw new ArgumentNullException(nameof(message));
        }

        /// <summary>Gets the stable diagnostic code.</summary>
        public string Code { get; }

        /// <summary>Gets diagnostic severity.</summary>
        public LegacyDocDiagnosticSeverity Severity { get; }

        /// <summary>Gets human-readable diagnostic text.</summary>
        public string Message { get; }

        /// <summary>
        /// Formats the diagnostic as a compact code/severity/message string.
        /// </summary>
        /// <returns>Formatted diagnostic text.</returns>
        public override string ToString() {
            return $"{Code} ({Severity}): {Message}";
        }
    }
}
