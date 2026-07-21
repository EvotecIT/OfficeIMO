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

    /// <summary>Controls whether a conversion may continue when content loss is known.</summary>
    public enum ExcelConversionLossPolicy {
        /// <summary>Reject conversion when known content would be omitted.</summary>
        Block,
        /// <summary>Continue and report known omitted content in the result.</summary>
        Allow
    }

    /// <summary>Controls how conversion handles an existing destination file.</summary>
    public enum ExcelConversionFileConflictPolicy {
        /// <summary>Reject conversion if the destination exists.</summary>
        FailIfExists,
        /// <summary>Replace an existing destination through an atomic commit.</summary>
        Replace
    }

    /// <summary>Identifies the purpose of a conversion diagnostic.</summary>
    public enum ExcelConversionDiagnosticCategory {
        /// <summary>Source format detection or extension findings.</summary>
        SourceFormat,
        /// <summary>Content that cannot survive conversion.</summary>
        DataLoss,
        /// <summary>Destination format or writer findings.</summary>
        DestinationFormat
    }

    /// <summary>Identifies the severity of a conversion diagnostic.</summary>
    public enum ExcelConversionDiagnosticSeverity {
        /// <summary>Informational finding.</summary>
        Information,
        /// <summary>Finding requiring user review.</summary>
        Warning,
        /// <summary>Finding that prevented conversion.</summary>
        Error
    }

    /// <summary>Describes a structured Excel conversion finding.</summary>
    public sealed class ExcelConversionDiagnostic {
        internal ExcelConversionDiagnostic(
            string code,
            ExcelConversionDiagnosticCategory category,
            ExcelConversionDiagnosticSeverity severity,
            string message,
            bool representsDataLoss,
            OfficeCompatibilityState? compatibilityState = null,
            OfficeCompatibilityImpact compatibilityImpact = OfficeCompatibilityImpact.None,
            string? sourceLocation = null,
            string? fallbackArtifact = null) {
            Code = code;
            Category = category;
            Severity = severity;
            Message = message;
            RepresentsDataLoss = representsDataLoss;
            CompatibilityState = compatibilityState ?? InferCompatibilityState(category, representsDataLoss);
            CompatibilityImpact = compatibilityImpact == OfficeCompatibilityImpact.None && representsDataLoss
                ? OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Carrier
                : compatibilityImpact;
            SourceLocation = sourceLocation;
            FallbackArtifact = fallbackArtifact;
        }

        /// <summary>Gets the stable diagnostic code.</summary>
        public string Code { get; }

        /// <summary>Gets the diagnostic category.</summary>
        public ExcelConversionDiagnosticCategory Category { get; }

        /// <summary>Gets the diagnostic severity.</summary>
        public ExcelConversionDiagnosticSeverity Severity { get; }

        /// <summary>Gets the human-readable diagnostic message.</summary>
        public string Message { get; }

        /// <summary>Gets whether the diagnostic describes content that will not survive conversion.</summary>
        public bool RepresentsDataLoss { get; }

        /// <summary>Gets the shared feature-level representation state.</summary>
        public OfficeCompatibilityState CompatibilityState { get; }

        /// <summary>Gets the fidelity dimensions affected by the finding.</summary>
        public OfficeCompatibilityImpact CompatibilityImpact { get; }

        /// <summary>Gets the related source part, sheet, range, record, or other location.</summary>
        public string? SourceLocation { get; }

        /// <summary>Gets the generated fallback artifact, when one exists.</summary>
        public string? FallbackArtifact { get; }

        private static OfficeCompatibilityState InferCompatibilityState(
            ExcelConversionDiagnosticCategory category,
            bool representsDataLoss) {
            if (category == ExcelConversionDiagnosticCategory.DestinationFormat) return OfficeCompatibilityState.Blocked;
            return representsDataLoss ? OfficeCompatibilityState.Dropped : OfficeCompatibilityState.Equivalent;
        }
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
            IReadOnlyList<ExcelConversionDiagnostic> diagnostics,
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
            IReadOnlyList<ExcelConversionDiagnostic> diagnostics,
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
        public IReadOnlyList<ExcelConversionDiagnostic> Diagnostics { get; }

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

        private static OfficeCompatibilityFinding CreateCompatibilityFinding(ExcelConversionDiagnostic diagnostic) {
            OfficeCompatibilityState state = diagnostic.CompatibilityState;
            OfficeCompatibilitySeverity severity = diagnostic.Severity switch {
                ExcelConversionDiagnosticSeverity.Warning => OfficeCompatibilitySeverity.Warning,
                ExcelConversionDiagnosticSeverity.Error => OfficeCompatibilitySeverity.Error,
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

    /// <summary>Identifies why an Excel conversion was rejected.</summary>
    public enum ExcelDocumentConversionFailureReason {
        /// <summary>Source and destination physical formats are identical.</summary>
        SameFormat,
        /// <summary>The destination exists and replacement was not allowed.</summary>
        DestinationExists,
        /// <summary>Known content loss was blocked by policy.</summary>
        DataLossBlocked,
        /// <summary>The destination writer cannot represent source content.</summary>
        DestinationFeatureUnsupported
    }

    /// <summary>Raised when a validated Excel conversion cannot be completed safely.</summary>
    public sealed class ExcelDocumentConversionException : InvalidOperationException {
        internal ExcelDocumentConversionException(
            ExcelDocumentConversionFailureReason reason,
            ExcelDocumentConversionResult result,
            string message,
            Exception? innerException = null)
            : base(message, innerException) {
            Reason = reason;
            Result = result;
        }

        /// <summary>Gets the structured failure reason.</summary>
        public ExcelDocumentConversionFailureReason Reason { get; }

        /// <summary>Gets the conversion assessment available when the operation was rejected.</summary>
        public ExcelDocumentConversionResult Result { get; }
    }

    /// <summary>Controls file-to-file Excel workbook conversion.</summary>
    public sealed class ExcelDocumentConversionOptions {
        /// <summary>Gets or sets how an existing destination is handled. The default is to fail.</summary>
        public ExcelConversionFileConflictPolicy FileConflictPolicy { get; set; } = ExcelConversionFileConflictPolicy.FailIfExists;

        /// <summary>Gets or sets how known conversion loss is handled. The default is to block it.</summary>
        public ExcelConversionLossPolicy LossPolicy { get; set; } = ExcelConversionLossPolicy.Block;

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
        public OpenSettings? OpenSettings { get; set; }

        /// <summary>
        /// Gets or sets optional legacy XLS import settings. Conversion always enables unsupported-content
        /// discovery so <see cref="LossPolicy"/> cannot be bypassed by suppressing import diagnostics.
        /// </summary>
        public LegacyXlsImportOptions? LegacyXlsImportOptions { get; set; }

        /// <summary>Gets or sets optional save settings for the destination file.</summary>
        public ExcelSaveOptions? SaveOptions { get; set; }
    }
}
