using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc {
    /// <summary>
    /// Contains the projected OfficeIMO document and the legacy DOC import report produced from the same parse.
    /// </summary>
    internal sealed class LegacyDocLoadResult : OfficeConversionReport, IDisposable {
        private readonly WordDocument? _document;
        private readonly Lazy<LegacyDocImportReport> _importReport;
        private readonly Lazy<LegacyDocImportSummary> _summary;
        private readonly Lazy<IReadOnlyList<OfficeDiagnostic>> _fidelityDiagnostics;

        internal LegacyDocLoadResult(WordDocument? document, LegacyDocDocument legacyDocument, Exception? projectionException = null) {
            _document = document;
            LegacyDocument = legacyDocument ?? throw new ArgumentNullException(nameof(legacyDocument));
            ProjectionException = projectionException;
            _importReport = new Lazy<LegacyDocImportReport>(() => LegacyDocument.CreateImportReport());
            _summary = new Lazy<LegacyDocImportSummary>(() => new LegacyDocImportSummary(this));
            _fidelityDiagnostics = new Lazy<IReadOnlyList<OfficeDiagnostic>>(CreateFidelityDiagnostics);
        }

        /// <summary>
        /// Gets the normal OfficeIMO Word document projected from supported legacy DOC content.
        /// </summary>
        public WordDocument Document => _document ?? throw new InvalidOperationException("No OfficeIMO Word document was projected from the legacy DOC file. Inspect LegacyDocument, NativeDiagnostics, ImportReport, and ProjectionException for import details.", ProjectionException);

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
        [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
        public LegacyDocDocument LegacyDocument { get; }

        /// <summary>Gets the advanced neutral parser model for forensic or corpus analysis.</summary>
        public LegacyDocDocument AdvancedDocument => LegacyDocument;

        /// <summary>
        /// Gets diagnostics produced while reading the legacy document.
        /// </summary>
        public IReadOnlyList<LegacyDocImportDiagnostic> NativeDiagnostics => LegacyDocument.Diagnostics;

        /// <summary>
        /// Gets unsupported or preserve-only features discovered while reading the legacy document.
        /// </summary>
        public IReadOnlyList<LegacyDocUnsupportedFeature> UnsupportedFeatures => LegacyDocument.UnsupportedFeatures;

        /// <summary>
        /// Gets preserve-only non-compound feature metadata discovered while importing the legacy document.
        /// </summary>
        public IReadOnlyList<LegacyDocPreservedFeature> PreservedFeatures => LegacyDocument.PreservedFeatures;

        /// <summary>
        /// Gets preserve-only compound storage discovered while importing the legacy document.
        /// </summary>
        public IReadOnlyList<LegacyDocCompoundFeature> CompoundFeatures => LegacyDocument.CompoundFeatures;

        /// <summary>
        /// Gets a compact import report for corpus baselines and preflight checks.
        /// </summary>
        [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
        public LegacyDocImportReport ImportReport => _importReport.Value;

        /// <summary>Gets a compact cached summary intended for normal application code.</summary>
        public LegacyDocImportSummary Summary => _summary.Value;

        /// <summary>Creates or returns the cached advanced corpus-grade import report.</summary>
        public LegacyDocImportReport CreateAdvancedImportReport() => _importReport.Value;

        /// <summary>
        /// Gets whether the legacy DOC import produced error diagnostics.
        /// </summary>
        public bool HasImportErrors => ProjectionException != null ||
            NativeDiagnostics.Any(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error);

        /// <summary>
        /// Gets whether conversion to DOCX would omit unsupported, preserved-only, or compound legacy content.
        /// </summary>
        public bool HasConversionLoss => HasLoss;

        /// <inheritdoc />
        protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
        private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => _fidelityDiagnostics.Value;

        /// <inheritdoc />
        public override bool HasLoss => FidelityDiagnostics.Any(diagnostic =>
            diagnostic.LossKind != OfficeConversionLossKind.None);

        /// <inheritdoc />
        protected override OfficeConversionException CreateLossException() => new OfficeConversionException("The legacy DOC import reported content loss. Inspect Diagnostics and the source import collections for details.", this);

        /// <summary>
        /// Throws when the legacy DOC import produced error diagnostics.
        /// </summary>
        public LegacyDocLoadResult EnsureNoImportErrors() {
            if (ProjectionException != null) {
                throw new InvalidOperationException(
                    "Legacy DOC content was parsed but could not be projected to an OfficeIMO document.",
                    ProjectionException);
            }
            if (HasImportErrors) {
                throw new InvalidOperationException("Legacy DOC import produced errors: " + string.Join("; ", NativeDiagnostics.Where(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error).Take(8).Select(diagnostic => diagnostic.ToString())));
            }

            return this;
        }

        /// <summary>Throws when conversion to DOCX would omit known legacy content.</summary>
        public LegacyDocLoadResult EnsureNoConversionLoss() {
            if (HasConversionLoss) {
                throw new InvalidOperationException("Legacy DOC import contains unsupported, preserved-only, or compound content that cannot be projected to DOCX without loss.");
            }

            return this;
        }

        /// <summary>
        /// Disposes the projected OfficeIMO document.
        /// </summary>
        public void Dispose() {
            _document?.Dispose();
        }

        private IReadOnlyList<OfficeDiagnostic> CreateFidelityDiagnostics() {
            var diagnostics = new List<OfficeDiagnostic>();
            if (ProjectionException != null) {
                diagnostics.Add(new global::OfficeIMO.OfficeDiagnostic(
                    "doc-projection-failed",
                    global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                    ProjectionException.Message,
                    ProjectionException.GetType().FullName,
                    "OfficeIMO.Word.LegacyDoc.Projection",
                    OfficeConversionLossKind.Failure));
            }
            diagnostics.AddRange(NativeDiagnostics.Select(diagnostic => new global::OfficeIMO.OfficeDiagnostic(
                diagnostic.Code,
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                diagnostic.Message,
                null,
                "OfficeIMO.Word.LegacyDoc.Reader",
                ClassifyLoss(diagnostic))));
            diagnostics.AddRange(UnsupportedFeatures.Select(feature => new global::OfficeIMO.OfficeDiagnostic(
                feature.Code,
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                feature.Description,
                feature.EntryPath ?? feature.DetailCode,
                "OfficeIMO.Word.LegacyDoc.Reader",
                OfficeConversionLossKind.Omission)));
            diagnostics.AddRange(PreservedFeatures.Select(feature => new global::OfficeIMO.OfficeDiagnostic(
                feature.Code,
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                feature.Description,
                feature.DetailCode,
                "OfficeIMO.Word.LegacyDoc.Reader",
                OfficeConversionLossKind.Omission)));
            diagnostics.AddRange(CompoundFeatures.Select(feature => new global::OfficeIMO.OfficeDiagnostic(
                feature.Code,
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                feature.Description,
                feature.EntryPath ?? feature.DetailCode,
                "OfficeIMO.Word.LegacyDoc.Reader",
                OfficeConversionLossKind.Omission)));
            return Array.AsReadOnly(diagnostics.ToArray());
        }

        internal static OfficeConversionLossKind ClassifyLoss(LegacyDocImportDiagnostic diagnostic) {
            if (diagnostic.Severity == LegacyDocDiagnosticSeverity.Error) return OfficeConversionLossKind.Failure;
            // These reader failures leave an entire source story or embedded item unprojected,
            // unlike format/style recovery warnings where the content is still represented.
            if (diagnostic.Code is "doc-footnote-plc-invalid" or "doc-endnote-plc-invalid"
                or "doc-comment-plc-invalid" or "doc-bookmark-plc-invalid"
                or "doc-picture-data-invalid" or "doc-plcfhdd-invalid"
                or "doc-ole-properties-unreadable")
                return OfficeConversionLossKind.Omission;
            return diagnostic.Severity == LegacyDocDiagnosticSeverity.Warning
                ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.None;
        }
    }
}
