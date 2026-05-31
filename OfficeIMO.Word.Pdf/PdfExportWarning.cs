namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Describes content that could not be faithfully mapped during PDF export.
    /// </summary>
    public sealed class PdfExportWarning {
        /// <summary>
        /// Stable warning code suitable for assertions and wrapper routing.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Document location or feature area where the warning was produced.
        /// </summary>
        public string Source { get; }

        /// <summary>
        /// Human-readable warning message.
        /// </summary>
        public string Message { get; }

        /// <summary>
        /// Creates a new export warning.
        /// </summary>
        /// <param name="code">Stable warning code.</param>
        /// <param name="source">Document location or feature area.</param>
        /// <param name="message">Human-readable warning message.</param>
        public PdfExportWarning(string code, string source, string message) {
            Code = code ?? throw new System.ArgumentNullException(nameof(code));
            Source = source ?? throw new System.ArgumentNullException(nameof(source));
            Message = message ?? throw new System.ArgumentNullException(nameof(message));
        }

        /// <summary>
        /// Returns a readable representation of the warning.
        /// </summary>
        public override string ToString() => Code + ": " + Message;
    }
}
