using OfficeIMO.ContentSafety;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using OfficeIMO.Word;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

namespace OfficeIMO;

/// <summary>Format-neutral document inspection limits and native handler settings.</summary>
public sealed class OfficeInspectOptions {
    /// <summary>Bounds extraction and retained inspection metadata. Native inspection is enabled by default.</summary>
    public ReaderOptions Reader { get; set; } = CreateReaderOptions();
    /// <summary>Settings owned by the selected native format handler.</summary>
    public ReaderAllOptions Formats { get; set; } = new();

    /// <summary>Includes format detection, native counts and encryption evidence without requiring content extraction.</summary>
    public bool IncludeSummary { get; set; }
    /// <summary>Includes the format's detached inspection snapshot with the summary.</summary>
    public bool Detailed { get; set; }
    /// <summary>Optional summary detection overrides. Null preserves the common Reader policy.</summary>
    public ReaderDetectionOptions? Detection { get; set; }
    /// <summary>Optional immutable handler registry for summary detection and otherwise unsupported native formats.</summary>
    public OfficeDocumentReader? DocumentReader { get; set; }
    /// <summary>Optional native presentation inspection selection.</summary>
    public OfficeIMO.PowerPoint.PowerPointInspectionOptions? PowerPointInspection { get; set; }
    /// <summary>Optional native workbook inspection settings.</summary>
    public OfficeIMO.Excel.ExcelReaderOptions? ExcelInspection { get; set; }

    /// <summary>Maximum buffered input for stream inspection, in bytes.</summary>
    public long MaximumInputBytes { get; set; } = 64L * 1024L * 1024L;

    /// <summary>Includes native content and metadata. Disable this to inspect only the selected reports without extracting document content.</summary>
    public bool IncludeContent { get; set; } = true;
    /// <summary>Returns PDF read/rewrite capability blockers, including for unsupported or malformed sources.</summary>
    public PdfLoadOptions? PdfPreflight { get; set; }

    /// <summary>Inspects concealed content using the native bounded input path when supplied.</summary>
    public OfficeContentSafetyOptions? ContentSafety { get; set; }
    /// <summary>Inspects source provenance without rewriting the document when supplied.</summary>
    public OfficeProvenanceOptions? Provenance { get; set; }
    /// <summary>Inspects OPC package signatures when supplied.</summary>
    public OfficePackageSignatureInspectionOptions? PackageSignatures { get; set; }
    /// <summary>Inspects VBA signature metadata when supplied.</summary>
    public OfficeVbaSignatureInspectionOptions? VbaSignatures { get; set; }
    /// <summary>Inspects Word macro project signature evidence when supplied.</summary>
    public WordMacroProjectSignatureInspectionOptions? WordMacroSignatures { get; set; }
    /// <summary>Optional cryptographic provider for Word macro signature evidence; no provider is selected implicitly.</summary>
    public IOfficeSecurityProvider? SecurityProvider { get; set; }

    private static ReaderOptions CreateReaderOptions() {
        ReaderOptions options = ReaderOptions.CreateSafeIngestion();
        options.Inspection = new ReaderInspectionOptions();
        options.DetectionMode = ReaderDetectionMode.PreferContent;
        return options;
    }
}
