namespace OfficeIMO.Excel {
    /// <summary>
    /// Optional behaviors applied during <see cref="ExcelDocument.Save(string, ExcelSaveOptions?)"/> and
    /// <see cref="ExcelDocument.SaveAsync(string, ExcelSaveOptions?, System.Threading.CancellationToken)"/> to strengthen
    /// robustness and CI validation.
    /// </summary>
    public sealed class ExcelSaveOptions {
        /// <summary>Gets or sets whether to open the saved file after a successful file commit.</summary>
        public bool OpenAfterSave { get; set; }

        /// <summary>
        /// When true, attempts to repair common defined-name issues (duplicates, out-of-range LocalSheetId, #REF!) before save.
        /// </summary>
        public bool SafeRepairDefinedNames { get; set; }

        /// <summary>
        /// When true, validates the saved package using <c>OpenXmlValidator</c> and throws on any errors.
        /// </summary>
        public bool ValidateOpenXml { get; set; }

        /// <summary>
        /// When true, performs a safety preflight on all worksheets before saving, removing empty containers
        /// (e.g., empty Hyperlinks/MergeCells), dropping orphaned drawing/header-footer references, and cleaning
        /// up invalid table references. This can prevent rare "Repaired Records" notices in Excel.
        /// </summary>
        public bool SafePreflight { get; set; }

        /// <summary>
        /// When true, disables direct fast package writers and always uses the standard save finalization path.
        /// </summary>
        public bool DisableFastPackageWriter { get; set; }

        /// <summary>
        /// When true, evaluates supported formulas and writes cached values before this save.
        /// Unsupported formulas are preserved for Excel-compatible applications to calculate.
        /// </summary>
        public bool EvaluateFormulasBeforeSave { get; set; }

        /// <summary>
        /// When true, removes cached formula results before this save. Ignored when
        /// <see cref="EvaluateFormulasBeforeSave"/> is true.
        /// </summary>
        public bool ClearCachedFormulaResultsBeforeSave { get; set; }

        /// <summary>
        /// When true, marks formulas dirty before this save so Excel-compatible applications recalculate on open.
        /// </summary>
        public bool MarkFormulasDirtyBeforeSave { get; set; }

        /// <summary>
        /// When true, writes workbook calculation properties requesting a full recalculation on open.
        /// </summary>
        public bool ForceFullCalculationOnOpen { get; set; }

        /// <summary>
        /// Controls saves of workbooks projected from legacy XLS files when known legacy-only
        /// content cannot be represented by the selected output format. The default blocks the save.
        /// </summary>
        public ExcelConversionLossPolicy LossPolicy { get; set; } = ExcelConversionLossPolicy.Block;

        /// <summary>
        /// Returns an options instance with all features disabled.
        /// </summary>
        public static ExcelSaveOptions None => new ExcelSaveOptions();

        internal ExcelSaveOptions WithLossPolicy(ExcelConversionLossPolicy lossPolicy) {
            return new ExcelSaveOptions {
                OpenAfterSave = OpenAfterSave,
                SafeRepairDefinedNames = SafeRepairDefinedNames,
                ValidateOpenXml = ValidateOpenXml,
                SafePreflight = SafePreflight,
                DisableFastPackageWriter = DisableFastPackageWriter,
                EvaluateFormulasBeforeSave = EvaluateFormulasBeforeSave,
                ClearCachedFormulaResultsBeforeSave = ClearCachedFormulaResultsBeforeSave,
                MarkFormulasDirtyBeforeSave = MarkFormulasDirtyBeforeSave,
                ForceFullCalculationOnOpen = ForceFullCalculationOnOpen,
                LossPolicy = lossPolicy
            };
        }
    }
}
