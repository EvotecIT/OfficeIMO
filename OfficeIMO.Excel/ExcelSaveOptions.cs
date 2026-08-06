namespace OfficeIMO.Excel {
    /// <summary>
    /// Optional behaviors applied during <see cref="ExcelDocument.Save(string, ExcelSaveOptions?)"/> and
    /// <see cref="ExcelDocument.SaveAsync(string, ExcelSaveOptions?, System.Threading.CancellationToken)"/> to strengthen
    /// robustness and CI validation.
    /// </summary>
    public sealed class ExcelSaveOptions {
        /// <summary>Default maximum package size materialized by save operations: 256 MiB.</summary>
        public const long DefaultMaxInMemoryPackageBytes = 256L * 1024L * 1024L;

        /// <summary>Default maximum temporary package size staged for a non-seekable destination: 256 MiB.</summary>
        public const long DefaultMaxTemporaryPackageBytes = 256L * 1024L * 1024L;

        /// <summary>
        /// Maximum package size that may be materialized in memory while saving. The default is
        /// 256 MiB. Set to <c>null</c> only for explicitly trusted, intentionally larger workbooks.
        /// </summary>
        public long? MaxInMemoryPackageBytes { get; set; } = DefaultMaxInMemoryPackageBytes;

        /// <summary>
        /// Maximum package size that may be staged in a temporary file when a target framework
        /// requires a seekable package destination. The default is 256 MiB. Set to <c>null</c>
        /// only for explicitly trusted, intentionally larger workbooks.
        /// </summary>
        public long? MaxTemporaryPackageBytes { get; set; } = DefaultMaxTemporaryPackageBytes;

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
        public OfficeConversionLossPolicy LossPolicy { get; set; } = OfficeConversionLossPolicy.Block;

        /// <summary>
        /// Gets or sets how save operations handle digital-signature metadata. The safe default blocks
        /// package rewriting; removing or preserving invalidated markup must be selected explicitly.
        /// </summary>
        public OfficeSignatureMutationPolicy SignatureMutationPolicy { get; set; } =
            OfficeSignatureMutationPolicy.BlockSave;

        /// <summary>Returns a fresh options instance with the default save policy.</summary>
        public static ExcelSaveOptions Default => new ExcelSaveOptions();

        internal ExcelSaveOptions WithLossPolicy(OfficeConversionLossPolicy lossPolicy) {
            if (MaxInMemoryPackageBytes.HasValue && MaxInMemoryPackageBytes.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(MaxInMemoryPackageBytes));
            }
            if (MaxTemporaryPackageBytes.HasValue && MaxTemporaryPackageBytes.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(MaxTemporaryPackageBytes));
            }
            return new ExcelSaveOptions {
                MaxInMemoryPackageBytes = MaxInMemoryPackageBytes,
                MaxTemporaryPackageBytes = MaxTemporaryPackageBytes,
                SafeRepairDefinedNames = SafeRepairDefinedNames,
                ValidateOpenXml = ValidateOpenXml,
                SafePreflight = SafePreflight,
                DisableFastPackageWriter = DisableFastPackageWriter,
                EvaluateFormulasBeforeSave = EvaluateFormulasBeforeSave,
                ClearCachedFormulaResultsBeforeSave = ClearCachedFormulaResultsBeforeSave,
                MarkFormulasDirtyBeforeSave = MarkFormulasDirtyBeforeSave,
                ForceFullCalculationOnOpen = ForceFullCalculationOnOpen,
                LossPolicy = lossPolicy,
                SignatureMutationPolicy = SignatureMutationPolicy
            };
        }
    }
}
