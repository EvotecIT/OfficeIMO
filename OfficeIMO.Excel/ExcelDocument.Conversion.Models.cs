using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.LegacyXls;

namespace OfficeIMO.Excel {
    /// <summary>Identifies an Excel file's physical format.</summary>
    public enum ExcelFileFormat {
        /// <summary>Office Open XML Excel workbook.</summary>
        Xlsx,
        /// <summary>Excel 97-2003 BIFF8 workbook.</summary>
        Xls
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
            bool representsDataLoss) {
            Code = code;
            Category = category;
            Severity = severity;
            Message = message;
            RepresentsDataLoss = representsDataLoss;
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
    }

    /// <summary>Reports the outcome and fidelity findings of an Excel file conversion.</summary>
    public sealed class ExcelDocumentConversionResult {
        internal ExcelDocumentConversionResult(
            string sourcePath,
            string destinationPath,
            ExcelFileFormat sourceFormat,
            ExcelFileFormat destinationFormat,
            IReadOnlyList<ExcelConversionDiagnostic> diagnostics,
            bool outputCreated,
            bool replacedExistingFile) {
            SourcePath = sourcePath;
            DestinationPath = destinationPath;
            SourceFormat = sourceFormat;
            DestinationFormat = destinationFormat;
            Diagnostics = diagnostics;
            OutputCreated = outputCreated;
            ReplacedExistingFile = replacedExistingFile;
            HasDataLoss = diagnostics.Any(diagnostic => diagnostic.RepresentsDataLoss);
        }

        /// <summary>Gets the normalized source path.</summary>
        public string SourcePath { get; }

        /// <summary>Gets the normalized destination path.</summary>
        public string DestinationPath { get; }

        /// <summary>Gets the source's detected physical format.</summary>
        public ExcelFileFormat SourceFormat { get; }

        /// <summary>Gets the requested destination physical format.</summary>
        public ExcelFileFormat DestinationFormat { get; }

        /// <summary>Gets compact conversion diagnostics.</summary>
        public IReadOnlyList<ExcelConversionDiagnostic> Diagnostics { get; }

        /// <summary>Gets whether the conversion intentionally omitted known source content.</summary>
        public bool HasDataLoss { get; }

        /// <summary>Gets whether the destination artifact was committed successfully.</summary>
        public bool OutputCreated { get; }

        /// <summary>Gets whether a pre-existing destination file was replaced.</summary>
        public bool ReplacedExistingFile { get; }
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

        /// <summary>Gets or sets whether to open the destination after a successful commit.</summary>
        public bool OpenAfterSave { get; set; }

        /// <summary>Gets or sets optional Open XML load settings for XLSX sources.</summary>
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
