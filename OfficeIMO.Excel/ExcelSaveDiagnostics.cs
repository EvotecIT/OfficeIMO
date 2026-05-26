namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes which package writer was used by the most recent save operation.
    /// </summary>
    public sealed class ExcelSaveDiagnostics {
        internal ExcelSaveDiagnostics(ExcelSavePackageWriter writer, string? fastPackageSkipReason) {
            Writer = writer;
            FastPackageSkipReason = fastPackageSkipReason;
        }

        /// <summary>
        /// The package writer used by the save operation.
        /// </summary>
        public ExcelSavePackageWriter Writer { get; }

        /// <summary>
        /// True when the save used a fast package path instead of full package finalization.
        /// </summary>
        public bool UsedFastPackageWriter => Writer == ExcelSavePackageWriter.SimplePackage
            || Writer == ExcelSavePackageWriter.ExtendedPackage
            || Writer == ExcelSavePackageWriter.UnchangedPackage
            || Writer == ExcelSavePackageWriter.DirectDataSetPackage;

        /// <summary>
        /// Reason the fast package writer was skipped, when the save fell back to full package finalization.
        /// </summary>
        public string? FastPackageSkipReason { get; }

        internal static ExcelSaveDiagnostics Standard(string? fastPackageSkipReason) =>
            new ExcelSaveDiagnostics(ExcelSavePackageWriter.StandardPackage, fastPackageSkipReason);

        internal static ExcelSaveDiagnostics SimplePackage() =>
            new ExcelSaveDiagnostics(ExcelSavePackageWriter.SimplePackage, fastPackageSkipReason: null);

        internal static ExcelSaveDiagnostics ExtendedPackage() =>
            new ExcelSaveDiagnostics(ExcelSavePackageWriter.ExtendedPackage, fastPackageSkipReason: null);

        internal static ExcelSaveDiagnostics DirectDataSetPackage() =>
            new ExcelSaveDiagnostics(ExcelSavePackageWriter.DirectDataSetPackage, fastPackageSkipReason: null);

        internal static ExcelSaveDiagnostics UnchangedPackage() =>
            new ExcelSaveDiagnostics(ExcelSavePackageWriter.UnchangedPackage, fastPackageSkipReason: null);
    }

    /// <summary>
    /// Identifies the package writer used by an Excel save operation.
    /// </summary>
    public enum ExcelSavePackageWriter {
        /// <summary>
        /// Full Open XML package finalization was used.
        /// </summary>
        StandardPackage,

        /// <summary>
        /// An unchanged package payload was copied directly.
        /// </summary>
        UnchangedPackage,

        /// <summary>
        /// A simple workbook package was written directly without full package finalization.
        /// </summary>
        SimplePackage,

        /// <summary>
        /// A workbook package with additional supported parts was written directly without full package finalization.
        /// </summary>
        ExtendedPackage,

        /// <summary>
        /// A DataSet import package was written directly from the retained DataSet export model.
        /// </summary>
        DirectDataSetPackage
    }
}
