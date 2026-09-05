using System.Text.Json.Serialization;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Workflows;

/// <summary>Controls publication behavior when one item in a redaction batch fails.</summary>
public enum PdfRedactionBatchPublicationPolicy {
    /// <summary>Prepare every item, then publish every artifact in one rollback-capable transaction.</summary>
    AtomicAll,
    /// <summary>Publish each successful item independently and continue after ordinary item failures.</summary>
    ContinuePerItem
}

/// <summary>Serializable, deterministic file-set request for a PDF redaction batch.</summary>
public sealed class PdfRedactionBatchRequest {
    /// <summary>Stable request schema.</summary>
    public const string CurrentSchema = "officeimo.pdf.redaction.batch-request.v1";
    /// <summary>Schema identifier.</summary>
    public string Schema { get; set; } = CurrentSchema;
    /// <summary>Workflow mode applied to every selected input.</summary>
    public PdfRedactionWorkflowMode Mode { get; set; }
    /// <summary>Root containing reviewed source PDFs.</summary>
    public required string InputRoot { get; set; }
    /// <summary>Optional relative input paths. When empty, <see cref="SearchPattern"/> is enumerated beneath <see cref="InputRoot"/>.</summary>
    public IList<string> InputPaths { get; set; } = new List<string>();
    /// <summary>File-system search pattern used when <see cref="InputPaths"/> is empty. Defaults to *.pdf.</summary>
    public string SearchPattern { get; set; } = "*.pdf";
    /// <summary>Whether pattern discovery recurses beneath the input root.</summary>
    public bool RecurseSubdirectories { get; set; } = true;
    /// <summary>Separate root for redacted outputs or existing outputs being verified.</summary>
    public string? OutputRoot { get; set; }
    /// <summary>Separate root for privacy-safe per-item evidence.</summary>
    public required string EvidenceRoot { get; set; }
    /// <summary>Separate root containing reviewed decision manifests for apply and verify modes.</summary>
    public string? DecisionsRoot { get; set; }
    /// <summary>Suffix appended to each input stem for output PDFs.</summary>
    public string OutputSuffix { get; set; } = ".redacted.pdf";
    /// <summary>Suffix appended to each input stem for per-item evidence.</summary>
    public string EvidenceSuffix { get; set; } = ".redaction.json";
    /// <summary>Suffix appended to each input stem when resolving reviewed decisions.</summary>
    public string DecisionsSuffix { get; set; } = ".decisions.json";
    /// <summary>Consolidated privacy-safe batch manifest path.</summary>
    public required string ManifestPath { get; set; }
    /// <summary>Publication behavior for item failures.</summary>
    public PdfRedactionBatchPublicationPolicy PublicationPolicy { get; set; } = PdfRedactionBatchPublicationPolicy.AtomicAll;
    /// <summary>Output conflict behavior shared by the batch.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; } = OfficeWorkflowConflictPolicy.Fail;
    /// <summary>Named redaction recipe shared by the batch.</summary>
    public required PdfRedactionRecipe Recipe { get; set; }
    /// <summary>Shared resource limits.</summary>
    public PdfRedactionWorkflowLimits Limits { get; set; } = new();
    /// <summary>Additional host inputs that publication may not replace.</summary>
    [JsonIgnore]
    public IList<string> ProtectedInputPaths { get; set; } = new List<string>();
    /// <summary>Runtime OCR engine; never serialized.</summary>
    [JsonIgnore]
    public IOcrEngine? OcrEngine { get; set; }
    /// <summary>Runtime OCR options; never serialized.</summary>
    [JsonIgnore]
    public PdfOcrMergeOptions? OcrOptions { get; set; }
    /// <summary>Runtime owner password; never serialized.</summary>
    [JsonIgnore]
    public string? OwnerPassword { get; set; }
    /// <summary>Runtime output encryption; passwords are never serialized.</summary>
    [JsonIgnore]
    public PdfStandardEncryptionOptions? OutputEncryption { get; set; }
    /// <summary>Runtime signer for signed derivatives; never serialized.</summary>
    [JsonIgnore]
    public IPdfExternalSigner? OutputSigner { get; set; }
    /// <summary>Runtime signature options; never serialized.</summary>
    [JsonIgnore]
    public PdfExternalSignatureOptions? OutputSignatureOptions { get; set; }
    /// <summary>Runtime signature validator; never serialized.</summary>
    [JsonIgnore]
    public IPdfSignatureCryptographyProvider? OutputSignatureValidator { get; set; }
    /// <summary>Runtime validators; never serialized.</summary>
    [JsonIgnore]
    public IList<IPdfRedactionCancellationAwareExternalValidator> ExternalValidators { get; set; } = new List<IPdfRedactionCancellationAwareExternalValidator>();
}

/// <summary>Privacy-safe persisted batch result without host paths or request identifiers.</summary>
public sealed class PdfRedactionBatchRecord {
    internal PdfRedactionBatchRecord(PdfRedactionBatchResult result) {
        Status = result.Status;
        PublishedAtomically = result.PublishedAtomically;
        Summary = result.Summary;
        Items = result.Items.Select(static item => new PdfRedactionWorkflowRecord(item)).ToArray();
    }

    /// <summary>Stable batch-result schema.</summary>
    public string Schema => PdfRedactionBatchResult.CurrentSchema;
    /// <summary>Aggregate batch status.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Whether all item artifacts were published atomically.</summary>
    public bool PublishedAtomically { get; }
    /// <summary>Privacy-safe aggregate summary.</summary>
    public string Summary { get; }
    /// <summary>Privacy-safe item records in deterministic input order.</summary>
    public IReadOnlyList<PdfRedactionWorkflowRecord> Items { get; }
}
