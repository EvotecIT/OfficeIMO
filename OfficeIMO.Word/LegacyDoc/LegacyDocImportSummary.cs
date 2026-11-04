namespace OfficeIMO.Word.LegacyDoc {
    /// <summary>Provides the compact, user-facing outcome of a legacy DOC import.</summary>
    internal sealed class LegacyDocImportSummary : OfficeConversionReport {
        internal LegacyDocImportSummary(LegacyDocLoadResult result) {
            ParagraphCount = result.LegacyDocument.Paragraphs.Count;
            DiagnosticCount = result.Diagnostics.Count;
            UnsupportedFeatureCount = result.UnsupportedFeatures.Count;
            PreservedFeatureCount = result.PreservedFeatures.Count;
            CompoundFeatureCount = result.CompoundFeatures.Count;
            HasImportErrors = result.HasImportErrors;
            FidelityDiagnostics = Array.AsReadOnly(result.Diagnostics.ToArray());
        }

        /// <summary>Gets the decoded paragraph count.</summary>
        public int ParagraphCount { get; }

        /// <summary>Gets the diagnostic count.</summary>
        public int DiagnosticCount { get; }

        /// <summary>Gets the unsupported feature count.</summary>
        public int UnsupportedFeatureCount { get; }

        /// <summary>Gets the preserved-only feature count.</summary>
        public int PreservedFeatureCount { get; }

        /// <summary>Gets the compound feature count.</summary>
        public int CompoundFeatureCount { get; }

        /// <summary>Gets whether import errors occurred.</summary>
        public bool HasImportErrors { get; }

        /// <summary>Gets whether DOCX conversion would omit known content.</summary>
        public bool HasConversionLoss => HasLoss;

        /// <inheritdoc />
        protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
        private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }

        /// <inheritdoc />
        public override bool HasLoss => FidelityDiagnostics.Any(diagnostic =>
            diagnostic.LossKind != OfficeConversionLossKind.None);

        /// <inheritdoc />
        protected override OfficeConversionException CreateLossException() => new OfficeConversionException("The legacy DOC import summary reported content loss. Inspect Diagnostics for details.", this);
    }
}
