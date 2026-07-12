namespace OfficeIMO.Excel.LegacyXls {
    /// <summary>Provides the compact, user-facing outcome of a legacy XLS import.</summary>
    public sealed class LegacyXlsImportSummary {
        internal LegacyXlsImportSummary(LegacyXlsLoadResult result) {
            WorksheetCount = result.Workbook.Worksheets.Count;
            ChartSheetCount = result.ChartSheets.Count;
            DiagnosticCount = result.Diagnostics.Count;
            UnsupportedFeatureCount = result.UnsupportedFeatures.Count;
            PreservedFeatureCount = result.PreservedFeatures.Count;
            UnsupportedSheetCount = result.UnsupportedSheets.Count;
            CompoundFeatureCount = result.CompoundFeatures.Count;
            HasImportErrors = result.HasImportErrors;
            HasConversionLoss = result.HasConversionLoss;
        }

        /// <summary>Gets the projected worksheet count.</summary>
        public int WorksheetCount { get; }

        /// <summary>Gets the projected chart-sheet count.</summary>
        public int ChartSheetCount { get; }

        /// <summary>Gets the diagnostic count.</summary>
        public int DiagnosticCount { get; }

        /// <summary>Gets the unsupported feature count.</summary>
        public int UnsupportedFeatureCount { get; }

        /// <summary>Gets the preserve-only BIFF feature count.</summary>
        public int PreservedFeatureCount { get; }

        /// <summary>Gets the unsupported sheet count.</summary>
        public int UnsupportedSheetCount { get; }

        /// <summary>Gets the compound feature count.</summary>
        public int CompoundFeatureCount { get; }

        /// <summary>Gets whether import errors occurred.</summary>
        public bool HasImportErrors { get; }

        /// <summary>Gets whether XLSX conversion would omit known content.</summary>
        public bool HasConversionLoss { get; }
    }
}
