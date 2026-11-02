using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    /// <summary>Contains the projected OfficeIMO presentation and its binary PPT import report.</summary>
    internal sealed class LegacyPptLoadResult : OfficeConversionReport, IDisposable {
        private readonly PowerPointDocument? _document;
        private readonly Lazy<LegacyPptImportReport> _report;
        private readonly Lazy<IReadOnlyList<OfficeDiagnostic>> _fidelityDiagnostics;

        internal LegacyPptLoadResult(PowerPointDocument? document, LegacyPptDocument presentation,
            Exception? projectionException = null) {
            _document = document;
            SourceDocument = presentation ?? throw new ArgumentNullException(nameof(presentation));
            ProjectionException = projectionException;
            _report = new Lazy<LegacyPptImportReport>(presentation.CreateImportReport);
            _fidelityDiagnostics = new Lazy<IReadOnlyList<OfficeDiagnostic>>(CreateFidelityDiagnostics);
        }

        /// <summary>Gets the normal editable OfficeIMO presentation projected from supported binary content.</summary>
        public PowerPointDocument Document => _document ?? throw new InvalidOperationException(
            "No OfficeIMO presentation was projected. Inspect SourceDocument, NativeDiagnostics, ImportReport, and ProjectionException.",
            ProjectionException);

        /// <summary>Gets whether projection produced an OfficeIMO presentation.</summary>
        public bool HasDocument => _document != null;

        /// <summary>Gets the neutral binary presentation model.</summary>
        public LegacyPptDocument SourceDocument { get; }

        /// <summary>Gets a projection failure captured for report callers.</summary>
        public Exception? ProjectionException { get; }

        /// <summary>Gets parser diagnostics.</summary>
        public IReadOnlyList<LegacyPptImportDiagnostic> NativeDiagnostics => SourceDocument.Diagnostics;

        /// <summary>Gets the cached import report.</summary>
        public LegacyPptImportReport ImportReport => _report.Value;

        /// <summary>Gets whether import produced error diagnostics.</summary>
        public bool HasImportErrors => ProjectionException != null ||
            NativeDiagnostics.Any(diagnostic => diagnostic.Severity == LegacyPptDiagnosticSeverity.Error);

        /// <summary>Gets whether projection is known to omit unsupported content.</summary>
        public bool HasConversionLoss => HasLoss;

        /// <inheritdoc />
        protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
        private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => _fidelityDiagnostics.Value;

        /// <inheritdoc />
        public override bool HasLoss => FidelityDiagnostics.Any(static diagnostic =>
            diagnostic.LossKind != OfficeConversionLossKind.None);

        /// <inheritdoc />
        protected override OfficeConversionException CreateLossException() => new OfficeConversionException("The legacy PPT import reported content loss. Inspect FidelityDiagnostics and the source import report for details.", this);

        /// <summary>Throws when import produced errors.</summary>
        public LegacyPptLoadResult EnsureNoImportErrors() {
            if (ProjectionException != null) {
                throw new InvalidOperationException(
                    "Legacy PPT content was parsed but could not be projected to an OfficeIMO presentation.",
                    ProjectionException);
            }
            if (HasImportErrors) throw new InvalidOperationException("Legacy PPT import produced errors: "
                + string.Join("; ", NativeDiagnostics.Where(diagnostic => diagnostic.Severity == LegacyPptDiagnosticSeverity.Error)
                    .Take(8)));
            return this;
        }

        /// <summary>Throws when conversion would omit known content.</summary>
        public LegacyPptLoadResult EnsureNoConversionLoss() {
            if (HasConversionLoss) throw new InvalidOperationException(
                "Legacy PPT import contains unsupported content that cannot be projected without loss.");
            return this;
        }

        /// <inheritdoc />
        public void Dispose() => _document?.Dispose();

        private IReadOnlyList<OfficeDiagnostic> CreateFidelityDiagnostics() {
            var diagnostics = new List<OfficeDiagnostic>(ImportReport.Diagnostics);
            if (ProjectionException != null) {
                diagnostics.Insert(0, new global::OfficeIMO.OfficeDiagnostic(
                    "ppt-projection-failed",
                    global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                    ProjectionException.Message,
                    ProjectionException.GetType().FullName,
                    "OfficeIMO.PowerPoint.LegacyPpt.Projection",
                    OfficeConversionLossKind.Failure));
            }
            return Array.AsReadOnly(diagnostics.ToArray());
        }
    }
}
