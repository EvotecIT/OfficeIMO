using System.Collections.ObjectModel;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

/// <summary>Operations exposed by the cross-format provenance workflow.</summary>
public enum OfficeProvenanceWorkflowOperation {
    /// <summary>Inspect standards-defined structural provenance carriers.</summary>
    Inspect,
    /// <summary>Combine structural, cryptographic, text-integrity, and provider-specific evidence.</summary>
    Assess,
    /// <summary>Remove selected standards-defined provenance carriers through the owning format API.</summary>
    Remove
}

/// <summary>One typed cross-format provenance request.</summary>
public sealed class OfficeProvenanceWorkflowRequest {
    /// <summary>Caller-provided identifier used by progress and batch results.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Requested provenance operation.</summary>
    public OfficeProvenanceWorkflowOperation Operation { get; set; }

    /// <summary>Input asset path.</summary>
    public required string InputPath { get; set; }

    /// <summary>Output asset path for removal. When omitted, a sibling provenance-cleaned name is used.</summary>
    public string? OutputPath { get; set; }

    /// <summary>Conflict behavior used when publishing a removal artifact.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; } = OfficeWorkflowConflictPolicy.Rename;

    /// <summary>Structural inspection limits used by <see cref="OfficeProvenanceWorkflowOperation.Inspect"/>.</summary>
    public OfficeProvenanceOptions Inspection { get; set; } = new();

    /// <summary>Combined assessment options used by <see cref="OfficeProvenanceWorkflowOperation.Assess"/>.</summary>
    public OfficeProvenanceAssessmentOptions Assessment { get; set; } = new();

    /// <summary>Selective removal policy used by <see cref="OfficeProvenanceWorkflowOperation.Remove"/>.</summary>
    public OfficeProvenanceRemovalOptions Removal { get; set; } = new();

    /// <summary>Shared input and output byte limits.</summary>
    public OfficeWorkflowLimits Limits { get; set; } = new();
}

/// <summary>Bounds sequential provenance batch execution.</summary>
public sealed class OfficeProvenanceWorkflowBatchOptions {
    /// <summary>Maximum number of materialized requests. Defaults to 256.</summary>
    public int MaximumRequests { get; set; } = 256;

    /// <summary>Whether execution continues after a failed request. Defaults to true.</summary>
    public bool ContinueOnFailure { get; set; } = true;

    internal OfficeProvenanceWorkflowBatchOptions CloneAndValidate() {
        if (MaximumRequests <= 0 || MaximumRequests > 10_000) {
            throw new ArgumentOutOfRangeException(nameof(MaximumRequests), "MaximumRequests must be between 1 and 10,000.");
        }
        return new OfficeProvenanceWorkflowBatchOptions {
            MaximumRequests = MaximumRequests,
            ContinueOnFailure = ContinueOnFailure
        };
    }
}

/// <summary>Runs reusable cross-format provenance workflows.</summary>
public interface IOfficeProvenanceWorkflowRunner {
    /// <summary>Runs one provenance request.</summary>
    Task<OfficeProvenanceWorkflowResult> RunProvenanceAsync(
        OfficeProvenanceWorkflowRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>Runs a bounded provenance batch sequentially.</summary>
    Task<IReadOnlyList<OfficeProvenanceWorkflowResult>> RunProvenanceBatchAsync(
        IEnumerable<OfficeProvenanceWorkflowRequest> requests,
        OfficeProvenanceWorkflowBatchOptions? options = null,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

/// <summary>One discoverable provenance workflow capability.</summary>
public sealed class OfficeProvenanceWorkflowCapability {
    internal OfficeProvenanceWorkflowCapability(
        string id,
        string label,
        IEnumerable<string> extensions,
        string ownerPackage,
        bool canRemove,
        string notes) {
        Id = id;
        Label = label;
        Extensions = Array.AsReadOnly(extensions.OrderBy(static item => item, StringComparer.Ordinal).ToArray());
        OwnerPackage = ownerPackage;
        CanRemove = canRemove;
        Notes = notes;
    }

    /// <summary>Stable capability identifier.</summary>
    public string Id { get; }
    /// <summary>User-facing format label.</summary>
    public string Label { get; }
    /// <summary>Recognized filename extensions.</summary>
    public IReadOnlyList<string> Extensions { get; }
    /// <summary>Package that owns format-specific inspection and mutation semantics.</summary>
    public string OwnerPackage { get; }
    /// <summary>Whether selected carriers can be removed.</summary>
    public bool CanRemove { get; }
    /// <summary>Important capability boundary.</summary>
    public string Notes { get; }
    /// <summary>Whether structural inspection is supported.</summary>
    public bool CanInspect => true;
    /// <summary>Whether combined assessment is supported.</summary>
    public bool CanAssess => true;
}

/// <summary>Canonical format and owner catalog used by provenance workflow consumers.</summary>
public static class OfficeProvenanceWorkflowCatalog {
    private static readonly IReadOnlyList<OfficeProvenanceWorkflowCapability> CapabilitiesValue =
        Array.AsReadOnly(new[] {
            new OfficeProvenanceWorkflowCapability("word-openxml", "Word Open XML", [".docm", ".docx", ".dotm", ".dotx"], "OfficeIMO.Word", true, "Package signatures block mutation unless removal is explicitly authorized."),
            new OfficeProvenanceWorkflowCapability("excel-package", "Excel workbook package", [".xlam", ".xlsb", ".xlsm", ".xlsx", ".xltm", ".xltx"], "OfficeIMO.Excel", true, "SpreadsheetML and XLSB package identity are validated before mutation."),
            new OfficeProvenanceWorkflowCapability("powerpoint-openxml", "PowerPoint Open XML", [".potm", ".potx", ".ppam", ".ppsm", ".ppsx", ".pptm", ".pptx"], "OfficeIMO.PowerPoint", true, "Package signatures block mutation unless removal is explicitly authorized."),
            new OfficeProvenanceWorkflowCapability("visio-openxml", "Visio Open XML", [".vsdm", ".vsdx", ".vssm", ".vssx", ".vstm", ".vstx"], "OfficeIMO.Visio", true, "Package signatures block mutation unless removal is explicitly authorized."),
            new OfficeProvenanceWorkflowCapability("open-document", "OpenDocument", [".odg", ".odp", ".ods", ".odt", ".otg", ".otp", ".ots", ".ott"], "OfficeIMO.OpenDocument", true, "Encrypted OpenDocument packages cannot be rewritten."),
            new OfficeProvenanceWorkflowCapability("epub", "EPUB", [".epub"], "OfficeIMO.Epub", true, "Package structure is validated before inspection or mutation."),
            new OfficeProvenanceWorkflowCapability("pdf", "PDF", [".pdf"], "OfficeIMO.Pdf", true, "Removal is limited to provenance associations supported by the PDF owner."),
            new OfficeProvenanceWorkflowCapability("html", "HTML", [".htm", ".html"], "OfficeIMO.Html", true, "External resources are not fetched during inspection."),
            new OfficeProvenanceWorkflowCapability("markdown", "Markdown", [".markdown", ".md"], "OfficeIMO.Markdown", true, "Original BOM-aware UTF encoding is preserved by file mutation."),
            new OfficeProvenanceWorkflowCapability("core-images", "Image provenance", [".gif", ".jpeg", ".jpg", ".png", ".svg", ".tif", ".tiff", ".webp"], "OfficeIMO.Core", true, "Inspection is signature-based; extensions are discovery hints."),
            new OfficeProvenanceWorkflowCapability("core-text", "Structured text", [".adoc", ".asciidoc", ".bat", ".c", ".cjs", ".cmd", ".cpp", ".cs", ".css", ".go", ".h", ".hpp", ".ini", ".java", ".js", ".json", ".lua", ".mjs", ".ps1", ".py", ".rb", ".rs", ".sh", ".sql", ".tex", ".toml", ".ts", ".txt", ".vb", ".xml", ".yaml", ".yml"], "OfficeIMO.Core", true, "Only standards-defined structured or wrapped text carriers are changed."),
            new OfficeProvenanceWorkflowCapability("core-detected", "Signature-detected asset", Array.Empty<string>(), "OfficeIMO.Core", true, "Unknown extensions may be inspected and removed when signature detection resolves to a supported non-generic format; generic ZIP mutation is not exposed.")
        });

    private static readonly IReadOnlyDictionary<string, OfficeProvenanceWorkflowCapability> ByExtension =
        new ReadOnlyDictionary<string, OfficeProvenanceWorkflowCapability>(
            CapabilitiesValue
                .SelectMany(capability => capability.Extensions.Select(extension => (extension, capability)))
                .ToDictionary(static item => item.extension, static item => item.capability, StringComparer.OrdinalIgnoreCase));

    /// <summary>All cross-format provenance capabilities in stable display order.</summary>
    public static IReadOnlyList<OfficeProvenanceWorkflowCapability> All => CapabilitiesValue;

    /// <summary>Finds the configured owner by filename extension.</summary>
    public static OfficeProvenanceWorkflowCapability? FindByPath(string? path) {
        string extension = Path.GetExtension(path ?? string.Empty);
        return ByExtension.TryGetValue(extension, out OfficeProvenanceWorkflowCapability? capability)
            ? capability
            : null;
    }
}
