using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc {
    /// <summary>
    /// Contains the projected OfficeIMO document and the legacy DOC import report produced from the same parse.
    /// </summary>
    public sealed class LegacyDocLoadResult : IDisposable {
        private readonly WordDocument? _document;

        internal LegacyDocLoadResult(WordDocument? document, LegacyDocDocument legacyDocument, Exception? projectionException = null) {
            _document = document;
            LegacyDocument = legacyDocument ?? throw new ArgumentNullException(nameof(legacyDocument));
            ProjectionException = projectionException;
        }

        /// <summary>
        /// Gets the normal OfficeIMO Word document projected from supported legacy DOC content.
        /// </summary>
        public WordDocument Document => _document ?? throw new InvalidOperationException("No OfficeIMO Word document was projected from the legacy DOC file. Inspect LegacyDocument, Diagnostics, ImportReport, and ProjectionException for import details.", ProjectionException);

        /// <summary>
        /// Gets whether supported legacy DOC content was projected into a normal OfficeIMO Word document.
        /// </summary>
        public bool HasDocument => _document != null;

        /// <summary>
        /// Gets the projection failure captured while preserving parser diagnostics for report callers.
        /// </summary>
        public Exception? ProjectionException { get; }

        /// <summary>
        /// Gets the neutral legacy DOC model produced by the parser.
        /// </summary>
        public LegacyDocDocument LegacyDocument { get; }

        /// <summary>
        /// Gets diagnostics produced while reading the legacy document.
        /// </summary>
        public IReadOnlyList<LegacyDocImportDiagnostic> Diagnostics => LegacyDocument.Diagnostics;

        /// <summary>
        /// Gets a compact import report for corpus baselines and preflight checks.
        /// </summary>
        public LegacyDocImportReport ImportReport => LegacyDocument.CreateImportReport();

        /// <summary>
        /// Gets whether the legacy DOC import produced error diagnostics.
        /// </summary>
        public bool HasImportErrors => Diagnostics.Any(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error);

        /// <summary>
        /// Throws when the legacy DOC import produced error diagnostics.
        /// </summary>
        public LegacyDocLoadResult EnsureNoImportErrors() {
            if (HasImportErrors) {
                throw new InvalidOperationException("Legacy DOC import produced errors: " + string.Join("; ", Diagnostics.Where(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error).Take(8).Select(diagnostic => diagnostic.ToString())));
            }

            return this;
        }

        /// <summary>
        /// Disposes the projected OfficeIMO document.
        /// </summary>
        public void Dispose() {
            _document?.Dispose();
        }
    }
}
