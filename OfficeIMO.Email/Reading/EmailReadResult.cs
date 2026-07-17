namespace OfficeIMO.Email;

/// <summary>Result of a bounded email artifact read.</summary>
public sealed class EmailReadResult : IDisposable {
    private readonly IDisposable? _resources;
    private bool _disposed;

    internal EmailReadResult(EmailDocument document, IReadOnlyList<EmailDiagnostic> diagnostics, long bytesRead,
        IDisposable? resources = null) {
        Document = document;
        Diagnostics = diagnostics;
        BytesRead = bytesRead;
        _resources = resources;
    }

    /// <summary>Parsed artifact document.</summary>
    public EmailDocument Document { get; }

    /// <summary>Structured compatibility and fidelity diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>True when parsing produced at least one error diagnostic.</summary>
    public bool HasErrors {
        get {
            foreach (EmailDiagnostic diagnostic in Diagnostics) {
                if (diagnostic.Severity == EmailDiagnosticSeverity.Error) return true;
            }
            return false;
        }
    }

    /// <summary>Number of source bytes consumed.</summary>
    public long BytesRead { get; }

    /// <summary>True when retained attachment payloads are reopenable file-backed sources.</summary>
    public bool UsesFileBackedContent => _resources is EmailReadWorkspace workspace && workspace.HasContent;

    /// <summary>Deletes temporary content owned by a streaming read result.</summary>
    public void Dispose() {
        if (_disposed) return;
        _resources?.Dispose();
        _disposed = true;
    }
}
