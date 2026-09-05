using System.Collections.ObjectModel;

namespace OfficeIMO.Workflows;

/// <summary>Operations supported by the local workflow runner.</summary>
public enum OfficeWorkflowOperation {
    /// <summary>Convert one supported document format to another.</summary>
    Convert,
    /// <summary>Inspect a PDF and return a health snapshot.</summary>
    Inspect,
    /// <summary>Compare two PDFs structurally and with the managed OfficeIMO renderer.</summary>
    Compare,
    /// <summary>Apply deterministic lossless PDF optimization.</summary>
    Optimize,
    /// <summary>Assess whether recovered PDF defects can be persisted as a verified repair artifact.</summary>
    RepairPlan,
    /// <summary>Create a verified normalized artifact from explicitly recovered PDF defects.</summary>
    Repair,
    /// <summary>Remove forbidden active content and embedded payloads from a PDF.</summary>
    Sanitize
}

/// <summary>Controls how an existing output path is handled.</summary>
public enum OfficeWorkflowConflictPolicy {
    /// <summary>Fail before publishing when the destination exists.</summary>
    Fail,
    /// <summary>Atomically replace the destination after the staged artifact is validated.</summary>
    Replace,
    /// <summary>Publish to the first available numbered file name.</summary>
    Rename
}

/// <summary>Named cross-format output intent.</summary>
public enum OfficeWorkflowOutputProfile {
    /// <summary>Preserve authored content and visual features where supported.</summary>
    Faithful,
    /// <summary>Prefer smaller output while retaining useful document structure.</summary>
    Lightweight,
    /// <summary>Prefer page setup and output intended for printing.</summary>
    PrintReady,
    /// <summary>Prefer text and tables over decorative visual content.</summary>
    TextOnly
}

/// <summary>Terminal state of one workflow request.</summary>
public enum OfficeWorkflowStatus {
    /// <summary>The request completed and any requested artifact was published.</summary>
    Completed,
    /// <summary>The request was cancelled cooperatively and no staged artifact was published.</summary>
    Cancelled,
    /// <summary>The request failed and no staged artifact was published.</summary>
    Failed
}

/// <summary>Stable category describing why a workflow did not complete.</summary>
public enum OfficeWorkflowFailureKind {
    /// <summary>The workflow completed or was cancelled without a failure.</summary>
    None,
    /// <summary>The request or one of its options was invalid.</summary>
    ValidationFailed,
    /// <summary>A requested input could not be found.</summary>
    InputNotFound,
    /// <summary>The input exists but is unsupported or structurally invalid.</summary>
    UnsupportedInput,
    /// <summary>The requested output could not be staged or published.</summary>
    OutputFailed,
    /// <summary>The workflow failed for another execution reason.</summary>
    OperationFailed
}

/// <summary>Severity of one structured workflow diagnostic.</summary>
public enum OfficeWorkflowDiagnosticSeverity {
    /// <summary>Informational evidence.</summary>
    Information,
    /// <summary>A fidelity limitation or review item.</summary>
    Warning,
    /// <summary>A failure that prevented completion.</summary>
    Error
}

/// <summary>Shared resource limits for one local document workflow.</summary>
public sealed class OfficeWorkflowLimits {
    /// <summary>Maximum primary or comparison input size.</summary>
    public long MaximumInputBytes { get; set; } = 256L * 1024L * 1024L;

    /// <summary>Maximum generated artifact size.</summary>
    public long MaximumOutputBytes { get; set; } = 512L * 1024L * 1024L;

    internal OfficeWorkflowLimits CloneAndValidate() {
        if (MaximumInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumInputBytes));
        if (MaximumOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumOutputBytes));
        return new OfficeWorkflowLimits {
            MaximumInputBytes = MaximumInputBytes,
            MaximumOutputBytes = MaximumOutputBytes
        };
    }
}

/// <summary>One immutable structured workflow diagnostic.</summary>
public sealed class OfficeWorkflowDiagnostic {
    /// <summary>Creates a workflow diagnostic.</summary>
    public OfficeWorkflowDiagnostic(
        string code,
        string message,
        OfficeWorkflowDiagnosticSeverity severity = OfficeWorkflowDiagnosticSeverity.Information,
        string? stage = null,
        IReadOnlyDictionary<string, string>? details = null) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Diagnostic code cannot be empty.", nameof(code));
        if (string.IsNullOrWhiteSpace(message)) throw new ArgumentException("Diagnostic message cannot be empty.", nameof(message));
        Code = code;
        Message = message;
        Severity = severity;
        Stage = stage;
        Details = new ReadOnlyDictionary<string, string>(
            new Dictionary<string, string>(details ?? new Dictionary<string, string>(), StringComparer.Ordinal));
    }

    /// <summary>Stable machine-readable code.</summary>
    public string Code { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Diagnostic severity.</summary>
    public OfficeWorkflowDiagnosticSeverity Severity { get; }
    /// <summary>Workflow stage that produced the diagnostic.</summary>
    public string? Stage { get; }
    /// <summary>Additional machine-readable evidence.</summary>
    public IReadOnlyDictionary<string, string> Details { get; }
}

/// <summary>Progress update for one workflow request.</summary>
public sealed class OfficeWorkflowProgress {
    /// <summary>Creates a progress update.</summary>
    public OfficeWorkflowProgress(string requestId, string stage, string message, double fraction, double? overallFraction = null) {
        RequestId = requestId;
        Stage = stage;
        Message = message;
        Fraction = Math.Clamp(fraction, 0D, 1D);
        OverallFraction = Math.Clamp(overallFraction ?? fraction, 0D, 1D);
    }

    /// <summary>Caller-provided request identifier.</summary>
    public string RequestId { get; }
    /// <summary>Stable execution stage.</summary>
    public string Stage { get; }
    /// <summary>User-facing progress message.</summary>
    public string Message { get; }
    /// <summary>Normalized completion fraction.</summary>
    public double Fraction { get; }
    /// <summary>Normalized overall batch fraction, equal to <see cref="Fraction"/> for a single request.</summary>
    public double OverallFraction { get; }
}

/// <summary>One typed local workflow request.</summary>
public sealed class OfficeWorkflowRequest {
    /// <summary>Caller-provided request identifier used by progress and batch results.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Requested operation.</summary>
    public OfficeWorkflowOperation Operation { get; set; }

    /// <summary>Primary input file.</summary>
    public required string InputPath { get; set; }

    /// <summary>Comparison input used by <see cref="OfficeWorkflowOperation.Compare"/>.</summary>
    public string? ComparisonPath { get; set; }

    /// <summary>Conversion route identifier from <see cref="OfficeWorkflowCatalog"/>.</summary>
    public string? ConversionRouteId { get; set; }

    /// <summary>Requested output file. Inspect does not require one; compare emits HTML when one is supplied.</summary>
    public string? OutputPath { get; set; }

    /// <summary>Conflict behavior used when publishing an artifact.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; } = OfficeWorkflowConflictPolicy.Rename;

    /// <summary>Cross-format output intent.</summary>
    public OfficeWorkflowOutputProfile OutputProfile { get; set; } = OfficeWorkflowOutputProfile.Faithful;

    /// <summary>Optional PDF password. It is used only while executing and is never copied to results or reports.</summary>
    public string? PdfPassword { get; set; }

    /// <summary>
    /// Optional password for the comparison PDF. When omitted, <see cref="PdfPassword"/> is reused.
    /// It is used only while executing and is never copied to results or reports.
    /// </summary>
    public string? ComparisonPdfPassword { get; set; }

    /// <summary>Shared request resource limits.</summary>
    public OfficeWorkflowLimits Limits { get; set; } = new();
}

/// <summary>Runs typed local OfficeIMO document workflows for desktop, command-line, and service hosts.</summary>
public interface IOfficeWorkflowRunner {
    /// <summary>Runs one workflow request.</summary>
    Task<OfficeWorkflowResult> RunAsync(
        OfficeWorkflowRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Snapshots and runs a bounded batch sequentially. Implementations reject more than
    /// <see cref="OfficeWorkflowRunner.MaximumBatchRequestCount"/> requests.
    /// </summary>
    Task<IReadOnlyList<OfficeWorkflowResult>> RunBatchAsync(
        IEnumerable<OfficeWorkflowRequest> requests,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

/// <summary>One supported conversion route projected from the canonical OfficeIMO capability catalog.</summary>
public sealed class OfficeWorkflowRoute {
    internal OfficeWorkflowRoute(OfficeConversionCapability capability, bool canExecute) {
        Id = capability.Id;
        Source = capability.Source;
        Target = capability.Target;
        SourceExtensions = Array.AsReadOnly(capability.SourceExtensions.ToArray());
        TargetExtension = capability.TargetExtension;
        Description = capability.Description;
        Fidelity = capability.Fidelity.ToString();
        SupportLevel = capability.SupportLevel.ToString();
        KnownLimitations = capability.KnownLimitations;
        Engine = capability.PackageId;
        Api = capability.Api;
        ResultContract = capability.ResultContract;
        InputKind = capability.InputKind;
        SupportEvidence = capability.SupportEvidence;
        TextFormatting = capability.TextFormatting;
        TextFormattingContract = capability.TextFormattingContract;
        BrowserAvailable = capability.BrowserAvailable;
        AgentDiscoverable = capability.AgentDiscoverable;
        CanExecute = canExecute;
    }

    /// <summary>Stable route identifier.</summary>
    public string Id { get; }
    /// <summary>Source format label.</summary>
    public string Source { get; }
    /// <summary>Destination format label.</summary>
    public string Target { get; }
    /// <summary>Accepted source extensions.</summary>
    public IReadOnlyList<string> SourceExtensions { get; }
    /// <summary>Destination extension.</summary>
    public string TargetExtension { get; }
    /// <summary>Capability description.</summary>
    public string Description { get; }
    /// <summary>Fidelity contract kind.</summary>
    public string Fidelity { get; }
    /// <summary>Evidence-based support level.</summary>
    public string SupportLevel { get; }
    /// <summary>Known route limitations.</summary>
    public string KnownLimitations { get; }
    /// <summary>First-party package that owns conversion semantics.</summary>
    public string Engine { get; }
    /// <summary>Representative public API owned by the route package.</summary>
    public string Api { get; }
    /// <summary>Public result type returned by the representative API.</summary>
    public string ResultContract { get; }
    /// <summary>How the route accepts its source.</summary>
    public OfficeConversionInputKind InputKind { get; }
    /// <summary>Evidence supporting the assigned support level.</summary>
    public string SupportEvidence { get; }
    /// <summary>Text and font formatting fidelity classification.</summary>
    public OfficeConversionTextFormattingKind TextFormatting { get; }
    /// <summary>Explicit text and font formatting promise.</summary>
    public string TextFormattingContract { get; }
    /// <summary>Whether the route is available in the shipped browser converter.</summary>
    public bool BrowserAvailable { get; }
    /// <summary>Whether the route is advertised through agent discovery.</summary>
    public bool AgentDiscoverable { get; }
    /// <summary>Whether this local workflow package can execute the route directly.</summary>
    public bool CanExecute { get; }
    /// <summary>User-facing route label.</summary>
    public string Label => Source + " to " + Target;
}

/// <summary>Canonical desktop/service conversion route view.</summary>
public static class OfficeWorkflowCatalog {
    private static readonly HashSet<string> ExecutableIds = new(StringComparer.Ordinal) {
        "docx-pdf", "xlsx-pdf", "pptx-pdf", "html-pdf",
        "pdf-docx", "pdf-xlsx", "pdf-pptx", "pdf-html"
    };

    private static readonly IReadOnlyList<OfficeWorkflowRoute> AllRoutesValue = Array.AsReadOnly(
        OfficeConversionCapabilityCatalog.All
            .Select(capability => new OfficeWorkflowRoute(capability, ExecutableIds.Contains(capability.Id)))
            .OrderBy(route => route.Source, StringComparer.Ordinal)
            .ThenBy(route => route.Target, StringComparer.Ordinal)
            .ToArray());

    private static readonly IReadOnlyList<OfficeWorkflowRoute> ExecutableRoutesValue = Array.AsReadOnly(
        AllRoutesValue.Where(static route => route.CanExecute).ToArray());

    /// <summary>All first-party conversion routes from the canonical capability catalog.</summary>
    public static IReadOnlyList<OfficeWorkflowRoute> Routes => AllRoutesValue;

    /// <summary>Routes this local workflow package can execute directly.</summary>
    public static IReadOnlyList<OfficeWorkflowRoute> ExecutableRoutes => ExecutableRoutesValue;

    /// <summary>Finds a catalog route by stable identifier.</summary>
    public static OfficeWorkflowRoute? Find(string? id) =>
        AllRoutesValue.FirstOrDefault(route => string.Equals(route.Id, id, StringComparison.OrdinalIgnoreCase));

    /// <summary>Finds a locally executable route by stable identifier.</summary>
    public static OfficeWorkflowRoute? FindExecutable(string? id) =>
        ExecutableRoutesValue.FirstOrDefault(route => string.Equals(route.Id, id, StringComparison.OrdinalIgnoreCase));

    /// <summary>Finds the unique catalog route matching source and target extensions.</summary>
    public static OfficeWorkflowRoute? Find(string sourceExtension, string targetExtension, bool executableOnly = false) {
        string source = NormalizeExtension(sourceExtension);
        string target = NormalizeExtension(targetExtension);
        IReadOnlyList<OfficeWorkflowRoute> routes = executableOnly ? ExecutableRoutesValue : AllRoutesValue;
        OfficeWorkflowRoute? match = null;
        foreach (OfficeWorkflowRoute route in routes) {
            if (!route.SourceExtensions.Contains(source, StringComparer.OrdinalIgnoreCase) ||
                !string.Equals(route.TargetExtension, target, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            // Some text-oriented routes intentionally share .txt. Ambiguous extension pairs
            // require an explicit route id instead of selecting a format by catalog order.
            if (match != null) return null;
            match = route;
        }

        return match;
    }

    private static string NormalizeExtension(string extension) {
        if (string.IsNullOrWhiteSpace(extension)) throw new ArgumentException("Extension cannot be empty.", nameof(extension));
        string value = extension.Trim();
        return (value.StartsWith(".", StringComparison.Ordinal) ? value : "." + value).ToLowerInvariant();
    }
}
