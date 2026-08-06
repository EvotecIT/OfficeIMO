using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.LegacyXls;

namespace OfficeIMO.Excel {
    /// <summary>Identifies an Excel file's physical format.</summary>
    public enum ExcelFileFormat {
        /// <summary>Office Open XML Excel workbook.</summary>
        Xlsx,
        /// <summary>Excel 97-2003 BIFF8 workbook.</summary>
        Xls,
        /// <summary>Excel Binary Workbook package containing BIFF12 record streams.</summary>
        Xlsb
    }

    /// <summary>Represents the destination artifact and report produced by an Excel file conversion.</summary>
    public sealed class ExcelDocumentConversionResult {
        internal ExcelDocumentConversionResult(
            string sourcePath,
            string destinationPath,
            ExcelFileFormat sourceFormat,
            ExcelFileFormat destinationFormat,
            OfficeFormatDescriptor sourceDescriptor,
            OfficeFormatDescriptor destinationDescriptor,
            IReadOnlyList<OfficeConversionDiagnostic> diagnostics,
            OfficeCompatibilityMode compatibilityMode,
            bool outputCreated,
            bool replacedExistingFile) {
            Value = outputCreated ? destinationPath : null;
            Report = new ExcelDocumentConversionReport(
                sourcePath,
                destinationPath,
                sourceFormat,
                destinationFormat,
                sourceDescriptor,
                destinationDescriptor,
                diagnostics,
                compatibilityMode,
                replacedExistingFile);
        }

        /// <summary>Gets the normalized destination path when the artifact was committed; otherwise, <see langword="null"/>.</summary>
        public string? Value { get; }

        /// <summary>Gets the immutable conversion assessment.</summary>
        public ExcelDocumentConversionReport Report { get; }

        /// <summary>Gets whether the conversion reported known content loss.</summary>
        public bool HasLoss => Report.HasLoss;

        /// <summary>Returns the committed destination path or throws when no artifact was produced.</summary>
        public string RequireValue() => Value
            ?? throw new InvalidOperationException("The Excel conversion did not produce a destination artifact.");

        /// <summary>Returns the committed destination path only when no content loss was reported.</summary>
        public string RequireNoLoss() {
            Report.RequireNoLoss();
            return RequireValue();
        }
    }

    /// <summary>Describes formats, paths, diagnostics, and commit behavior for one Excel conversion.</summary>
    public sealed class ExcelDocumentConversionReport {
        internal ExcelDocumentConversionReport(
            string sourcePath,
            string destinationPath,
            ExcelFileFormat sourceFormat,
            ExcelFileFormat destinationFormat,
            OfficeFormatDescriptor sourceDescriptor,
            OfficeFormatDescriptor destinationDescriptor,
            IReadOnlyList<OfficeConversionDiagnostic> diagnostics,
            OfficeCompatibilityMode compatibilityMode,
            bool replacedExistingFile) {
            SourcePath = sourcePath;
            DestinationPath = destinationPath;
            SourceFormat = sourceFormat;
            DestinationFormat = destinationFormat;
            SourceFormatDescriptor = sourceDescriptor;
            DestinationFormatDescriptor = destinationDescriptor;
            Diagnostics = Array.AsReadOnly((diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToArray());
            Compatibility = new OfficeCompatibilityReport(
                sourceDescriptor,
                destinationDescriptor,
                compatibilityMode,
                Diagnostics.Select(CreateCompatibilityFinding));
            ReplacedExistingFile = replacedExistingFile;
        }

        /// <summary>Gets the normalized source path.</summary>
        public string SourcePath { get; }

        /// <summary>Gets the normalized destination path.</summary>
        public string DestinationPath { get; }

        /// <summary>Gets the source's detected physical format.</summary>
        public ExcelFileFormat SourceFormat { get; }

        /// <summary>Gets the requested destination physical format.</summary>
        public ExcelFileFormat DestinationFormat { get; }

        /// <summary>Gets the concrete source format and document kind.</summary>
        public OfficeFormatDescriptor SourceFormatDescriptor { get; }

        /// <summary>Gets the concrete destination format and document kind.</summary>
        public OfficeFormatDescriptor DestinationFormatDescriptor { get; }

        /// <summary>Gets a snapshot of conversion diagnostics.</summary>
        public IReadOnlyList<OfficeConversionDiagnostic> Diagnostics { get; }

        /// <summary>Gets the shared feature-level fidelity assessment for this conversion.</summary>
        public OfficeCompatibilityReport Compatibility { get; }

        /// <summary>Gets whether the conversion reported known content loss.</summary>
        public bool HasLoss => Diagnostics.Any(static diagnostic => diagnostic.RepresentsDataLoss);

        /// <summary>Gets whether a pre-existing destination file was replaced.</summary>
        public bool ReplacedExistingFile { get; }

        /// <summary>Throws when the conversion reported known content loss.</summary>
        public void RequireNoLoss() {
            Compatibility.RequireNoLoss();
        }

        private static OfficeCompatibilityFinding CreateCompatibilityFinding(OfficeConversionDiagnostic diagnostic) {
            OfficeCompatibilityState state = diagnostic.CompatibilityState;
            OfficeCompatibilitySeverity severity = diagnostic.Severity switch {
                OfficeConversionDiagnosticSeverity.Warning => OfficeCompatibilitySeverity.Warning,
                OfficeConversionDiagnosticSeverity.Error => OfficeCompatibilitySeverity.Error,
                _ => OfficeCompatibilitySeverity.Information
            };
            return new OfficeCompatibilityFinding(
                diagnostic.Code,
                diagnostic.Category.ToString(),
                diagnostic.Message,
                state,
                severity,
                diagnostic.CompatibilityImpact,
                diagnostic.RepresentsDataLoss,
                diagnostic.SourceLocation,
                diagnostic.FallbackArtifact);
        }
    }

    /// <summary>Raised when a validated Excel conversion cannot be completed safely.</summary>
    public sealed class ExcelDocumentConversionException : InvalidOperationException {
        internal ExcelDocumentConversionException(
            OfficeConversionFailureReason reason,
            ExcelDocumentConversionResult result,
            string message,
            Exception? innerException = null)
            : base(message, innerException) {
            Reason = reason;
            Result = result;
        }

        /// <summary>Gets the structured failure reason.</summary>
        public OfficeConversionFailureReason Reason { get; }

        /// <summary>Gets the conversion assessment available when the operation was rejected.</summary>
        public ExcelDocumentConversionResult Result { get; }
    }

    /// <summary>Controls file-to-file Excel workbook conversion.</summary>
    public sealed class ExcelDocumentConversionOptions {
        /// <summary>Gets or sets how an existing destination is handled. The default is to fail.</summary>
        public OfficeConversionFileConflictPolicy FileConflictPolicy { get; set; } = OfficeConversionFileConflictPolicy.FailIfExists;

        /// <summary>Gets or sets how known conversion loss is handled. The default is to block it.</summary>
        public OfficeConversionLossPolicy LossPolicy { get; set; } = OfficeConversionLossPolicy.Block;

        /// <summary>
        /// Gets or sets the requested fidelity strategy. Existing <see cref="LossPolicy"/> behavior remains
        /// authoritative until a format-specific fallback is selected by the conversion planner.
        /// </summary>
        public OfficeCompatibilityMode CompatibilityMode { get; set; } = OfficeCompatibilityMode.StrictNative;

        /// <summary>
        /// Gets or sets whether a lossy conversion retains the complete original source in an inert,
        /// hash-verified OfficeIMO carrier. The source may contain macros, hidden content, or embedded
        /// payloads, so this is disabled unless explicitly requested or preservation-only mode is selected.
        /// </summary>
        public bool EmbedSourceWhenLossy { get; set; }

        /// <summary>Gets or sets the maximum cell columns used by the XLS/XLSB visual fallback. The default is 128.</summary>
        public int VisualFallbackMaxColumns { get; set; } = 128;

        /// <summary>Gets or sets the maximum cell rows used by the XLS/XLSB visual fallback. The default is 1024.</summary>
        public int VisualFallbackMaxRows { get; set; } = 1024;

        /// <summary>
        /// Gets or sets optional Open XML load settings for XLSX sources. Conversion always disables
        /// <see cref="OpenSettings.AutoSave"/> so source files are never modified as a load side effect.
        /// </summary>
        public OfficeOpenXmlLoadSettings? OpenSettings { get; set; }

        /// <summary>
        /// Gets or sets optional legacy XLS import settings. Conversion always enables unsupported-content
        /// discovery so <see cref="LossPolicy"/> cannot be bypassed by suppressing import diagnostics.
        /// </summary>
        public LegacyXlsImportOptions? LegacyXlsImportOptions { get; set; }

        /// <summary>Gets or sets optional save settings for the destination file.</summary>
        public ExcelSaveOptions? SaveOptions { get; set; }
    }
}
