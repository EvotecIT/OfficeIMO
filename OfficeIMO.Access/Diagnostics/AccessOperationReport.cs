namespace OfficeIMO.Access;

/// <summary>Disposition of a single operation for the assessed model and target.</summary>
public enum AccessOperationStatus {
    /// <summary>The operation has qualified support.</summary>
    Supported,
    /// <summary>A codec or required qualification is unavailable.</summary>
    Unsupported
}

/// <summary>A stable diagnostic code and explanation. No native content or credentials are included.</summary>
public sealed class AccessDiagnostic {
    internal AccessDiagnostic(string code, string message, Guid? objectId = null) { Code = code; Message = message; ObjectId = objectId; }
    /// <summary>Machine-readable code.</summary>
    public string Code { get; }
    /// <summary>Explanation of the evidence or missing support.</summary>
    public string Message { get; }
    /// <summary>Object identity when the diagnostic concerns a modeled object.</summary>
    public Guid? ObjectId { get; }
}

/// <summary>Immutable save assessment bound to one document revision.</summary>
public sealed class AccessOperationReport {
    internal AccessOperationReport(Guid documentId, long revision, AccessFileFormat target, AccessFormatProfile profile, IReadOnlyList<AccessDiagnostic> diagnostics, AccessOperationStatus status = AccessOperationStatus.Unsupported) {
        DocumentId = documentId; Revision = revision; TargetFormat = target; TargetProfile = profile; Diagnostics = diagnostics; Status = status;
    }
    /// <summary>Identity of the assessed document.</summary>
    public Guid DocumentId { get; }
    /// <summary>Revision assessed before output.</summary>
    public long Revision { get; }
    /// <summary>Selected output family.</summary>
    public AccessFileFormat TargetFormat { get; }
    /// <summary>Selected physical output generation.</summary>
    public AccessFormatProfile TargetProfile { get; }
    /// <summary>Support for this operation; lack of a writer is never classified as loss-free success.</summary>
    public AccessOperationStatus Status { get; }
    /// <summary>Reasons and qualification gaps.</summary>
    public IReadOnlyList<AccessDiagnostic> Diagnostics { get; }
    /// <summary>Rejects a report from another document or a stale revision.</summary>
    public void RequireCurrent(AccessDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        document.EnsureNotDisposed();
        if (document.Id != DocumentId || document.Revision != Revision || document.HasActiveUpdate) throw new InvalidOperationException("The Access assessment is stale or belongs to another document. Assess again after committing edits.");
    }
    /// <summary>Requires a supported operation without known loss. Unsupported codecs fail before output.</summary>
    public void RequireNoLoss() {
        if (Status != AccessOperationStatus.Supported) throw new AccessOperationNotSupportedException(this);
    }
}

/// <summary>An unavailable native operation with its machine-readable assessment.</summary>
public sealed class AccessOperationNotSupportedException : NotSupportedException {
    internal AccessOperationNotSupportedException(AccessOperationReport report) : base(string.Join(" ", report.Diagnostics.Select(d => d.Message))) { Report = report; }
    /// <summary>The assessment that prevented output.</summary>
    public AccessOperationReport Report { get; }
}
