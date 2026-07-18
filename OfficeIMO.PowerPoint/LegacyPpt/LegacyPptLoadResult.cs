using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    /// <summary>Contains the projected OfficeIMO presentation and its binary PPT import report.</summary>
    public sealed class LegacyPptLoadResult : IDisposable {
        private readonly PowerPointPresentation? _document;
        private readonly Lazy<LegacyPptImportReport> _report;

        internal LegacyPptLoadResult(PowerPointPresentation? document, LegacyPptPresentation presentation,
            Exception? projectionException = null) {
            _document = document;
            Presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
            ProjectionException = projectionException;
            _report = new Lazy<LegacyPptImportReport>(presentation.CreateImportReport);
        }

        /// <summary>Gets the normal editable OfficeIMO presentation projected from supported binary content.</summary>
        public PowerPointPresentation Document => _document ?? throw new InvalidOperationException(
            "No OfficeIMO presentation was projected. Inspect Presentation, Diagnostics, ImportReport, and ProjectionException.",
            ProjectionException);

        /// <summary>Gets whether projection produced an OfficeIMO presentation.</summary>
        public bool HasDocument => _document != null;

        /// <summary>Gets the neutral binary presentation model.</summary>
        public LegacyPptPresentation Presentation { get; }

        /// <summary>Gets a projection failure captured for report callers.</summary>
        public Exception? ProjectionException { get; }

        /// <summary>Gets parser diagnostics.</summary>
        public IReadOnlyList<LegacyPptImportDiagnostic> Diagnostics => Presentation.Diagnostics;

        /// <summary>Gets the cached import report.</summary>
        public LegacyPptImportReport ImportReport => _report.Value;

        /// <summary>Gets whether import produced error diagnostics.</summary>
        public bool HasImportErrors => Diagnostics.Any(diagnostic => diagnostic.Severity == LegacyPptDiagnosticSeverity.Error);

        /// <summary>Gets whether projection is known to omit unsupported content.</summary>
        public bool HasConversionLoss => ImportReport.HasConversionLoss;

        /// <summary>Throws when import produced errors.</summary>
        public LegacyPptLoadResult EnsureNoImportErrors() {
            if (HasImportErrors) throw new InvalidOperationException("Legacy PPT import produced errors: "
                + string.Join("; ", Diagnostics.Where(diagnostic => diagnostic.Severity == LegacyPptDiagnosticSeverity.Error)
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
    }
}
