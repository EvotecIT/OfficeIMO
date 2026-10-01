using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Invoicing.Validation;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Invoice input acquisition and publication settings for local files or permission-scoped storage providers.</summary>
public sealed record OfficeInvoiceStorageWorkflowRequest {
    /// <summary>Source location. Provider locations require InputStream.</summary>
    public string InputPath { get; init; } = string.Empty;
    /// <summary>Selected invoice operation.</summary>
    public OfficeInvoiceWorkflowOperation Operation { get; init; }
    /// <summary>Explicit output/inspection contract. Source editing retains its original contract.</summary>
    public InvoiceXmlOptions? Target { get; init; }
    /// <summary>Explicit exact-byte standards release, if required.</summary>
    public InvoiceSpecificationRelease? ValidationRelease { get; init; }
    /// <summary>Captured source-field replacements for EditSource.</summary>
    public InvoiceSourceEdits? SourceEdits { get; init; }
    /// <summary>Output location required only for writing operations.</summary>
    public string? OutputPath { get; init; }
    /// <summary>Optional reopenable permission-scoped input. The owner closes its streams and verifies the source before publication.</summary>
    public OfficeWorkflowStreamInput? InputStream { get; init; }
    /// <summary>Optional confirmed direct-write provider output with verified contents and durable recovery.</summary>
    public OfficeWorkflowStreamOutput? OutputStream { get; init; }
    /// <summary>Local collision policy. Provider outputs require Replace after the host obtains direct-write consent.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; init; } = OfficeWorkflowConflictPolicy.Fail;
    /// <summary>Host authorization for output locations, including active-document and recovery protection.</summary>
    public IOfficeWorkflowPublicationGuard? PublicationGuard { get; init; }
    /// <summary>Input capture limit, at most 16 MiB.</summary>
    public long MaximumInputBytes { get; init; } = InvoiceProfileDeclaration.MaximumXmlBytes;
    /// <summary>Returned and published artifact limit; defaults to 64 MiB.</summary>
    public long MaximumOutputBytes { get; init; } = 64L * 1024 * 1024;
    /// <summary>Rendering layout, cloned before input acquisition.</summary>
    public InvoicePdfLayoutOptions Layout { get; init; } = new();
    /// <summary>PDF settings, cloned before input acquisition.</summary>
    public PdfOptions? PdfOptions { get; init; }
}

/// <summary>Shared invoice execution over local or provider storage.</summary>
public interface IOfficeInvoiceWorkflowRunner {
    /// <summary>Captures bounded input, executes the invoice owner, and publishes successful output with source checks and recovery.</summary>
    Task<OfficeInvoiceStorageWorkflowResult> RunInvoiceAsync(OfficeInvoiceStorageWorkflowRequest request,
        InvoiceValidator? validator = null, CancellationToken cancellationToken = default);
}

/// <summary>Invoice execution and storage-publication evidence, including uncertain provider outcomes.</summary>
public sealed class OfficeInvoiceStorageWorkflowResult {
    internal OfficeInvoiceStorageWorkflowResult(OfficeWorkflowStatus status, OfficeInvoiceWorkflowResult? workflow,
        string? outputPath, string summary, List<OfficeWorkflowDiagnostic> diagnostics, OfficeWorkflowOutputRecovery? recovery = null) {
        Status = status; Workflow = workflow; OutputPath = outputPath; Summary = summary;
        Diagnostics = diagnostics.AsReadOnly(); Recovery = recovery;
    }
    /// <summary>Storage execution status. Unconfirmed requires checking the provider output and retained recovery.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Invoice owner report, or null when input acquisition failed.</summary>
    public OfficeInvoiceWorkflowResult? Workflow { get; }
    /// <summary>True only when invoice execution and requested publication completed.</summary>
    public bool Succeeded => Status == OfficeWorkflowStatus.Completed && Workflow?.Succeeded == true;
    /// <summary>Published output location, if confirmed.</summary>
    public string? OutputPath { get; }
    /// <summary>Execution or publication summary.</summary>
    public string Summary { get; }
    /// <summary>Storage and execution diagnostics; invoice findings remain in Workflow.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
    /// <summary>Durable output recovery retained after uncertain provider publication or failed cleanup.</summary>
    public OfficeWorkflowOutputRecovery? Recovery { get; }
}
