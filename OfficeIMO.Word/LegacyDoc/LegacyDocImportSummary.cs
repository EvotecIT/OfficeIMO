namespace OfficeIMO.Word.LegacyDoc {
    /// <summary>Provides the compact, user-facing outcome of a legacy DOC import.</summary>
    public sealed class LegacyDocImportSummary {
        internal LegacyDocImportSummary(LegacyDocLoadResult result) {
            ParagraphCount = result.LegacyDocument.Paragraphs.Count;
            DiagnosticCount = result.Diagnostics.Count;
            UnsupportedFeatureCount = result.UnsupportedFeatures.Count;
            PreservedFeatureCount = result.PreservedFeatures.Count;
            CompoundFeatureCount = result.CompoundFeatures.Count;
            HasImportErrors = result.HasImportErrors;
            HasConversionLoss = result.HasConversionLoss;
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
        public bool HasConversionLoss { get; }
    }
}
