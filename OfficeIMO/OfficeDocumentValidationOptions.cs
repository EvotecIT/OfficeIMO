using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO;

/// <summary>Native policies consumed by OfficeDocument.Validate. No setting enables repair or saving.</summary>
public sealed class OfficeDocumentValidationOptions {
    /// <summary>Package limits and active-content policy.</summary>
    public OfficePackageSecurityOptions? Security { get; set; }
    /// <summary>Rejects active, embedded and external package content using UntrustedDefaults, in addition to any stricter supplied limits.</summary>
    public bool Untrusted { get; set; }
    /// <summary>Password used only when loading a path input.</summary>
    public string? Password { get; set; }
    /// <summary>Workbook doctor and capability settings.</summary>
    public OfficeExcelValidationOptions? Excel { get; set; }
    /// <summary>PDF parser and permission policy; copied before a command password is applied.</summary>
    public PdfLoadOptions? PdfRead { get; set; }
    /// <summary>PDF compliance, preflight capabilities and rewrite comparison.</summary>
    public OfficePdfValidationOptions? Pdf { get; set; }
    /// <summary>Presentation accessibility policy.</summary>
    public PowerPointAccessibilityOptions? PowerPointAccessibility { get; set; }
    /// <summary>Document accessibility policy for Word.</summary>
    public WordAccessibilityOptions? WordAccessibility { get; set; }
    /// <summary>Workbook accessibility policy.</summary>
    public ExcelAccessibilityOptions? ExcelAccessibility { get; set; }
}

/// <summary>Read-only workbook checks, retaining native feature and repair-hint evidence.</summary>
public sealed class OfficeExcelValidationOptions {
    /// <summary>Doctor policy. RepairDefinedNames is rejected by validation.</summary>
    public ExcelWorkbookDoctorOptions Doctor { get; set; } = new();
    /// <summary>Capabilities to check; null selects every capability when this options object is supplied.</summary>
    public ExcelPreflightCapability[]? Capability { get; set; }
    /// <summary>Retains the native feature report in the common report.</summary>
    public bool IncludeFeatures { get; set; }
    /// <summary>Retains native repair hints as evidence without applying them.</summary>
    public bool IncludeRepairHints { get; set; }
}

/// <summary>PDF checks performed without rewriting either document.</summary>
public sealed class OfficePdfValidationOptions {
    /// <summary>Optional formal profile; null uses the document's configured profile.</summary>
    public PdfComplianceProfile? Profile { get; set; }
    /// <summary>Combines compliance readiness with exact-artifact external validation evidence.</summary>
    public bool Proof { get; set; }
    /// <summary>External results already obtained by the caller; validation does not launch executables.</summary>
    public PdfExternalValidationResult[]? ExternalValidation { get; set; }
    /// <summary>Optional read/rewrite capabilities to require during Structure checks.</summary>
    public PdfPreflightCapability[]? Capability { get; set; }
    /// <summary>Caller-owned rewritten document to compare during Structure checks.</summary>
    public PdfDocument? DifferenceDocument { get; set; }
    /// <summary>Rewritten PDF path to compare during Structure checks; exclusive with DifferenceDocument.</summary>
    public string? DifferencePath { get; set; }
    /// <summary>Native preservation signals and independent original/rewritten read policies.</summary>
    public PdfRewritePreservationOptions? Rewrite { get; set; }
}
