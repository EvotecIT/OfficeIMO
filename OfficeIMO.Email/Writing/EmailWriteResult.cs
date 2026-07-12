namespace OfficeIMO.Email;

/// <summary>Result of email artifact serialization.</summary>
public sealed class EmailWriteResult {
    internal EmailWriteResult(long bytesWritten, IReadOnlyList<EmailDiagnostic> diagnostics, bool usedPreservedSource) {
        BytesWritten = bytesWritten;
        Diagnostics = diagnostics;
        UsedPreservedSource = usedPreservedSource;
    }

    /// <summary>Number of bytes written.</summary>
    public long BytesWritten { get; }

    /// <summary>Structured fidelity diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>True when serialization produced at least one error diagnostic.</summary>
    public bool HasErrors {
        get {
            foreach (EmailDiagnostic diagnostic in Diagnostics) {
                if (diagnostic.Severity == EmailDiagnosticSeverity.Error) return true;
            }
            return false;
        }
    }

    /// <summary>True when the original preserved bytes were emitted verbatim.</summary>
    public bool UsedPreservedSource { get; }
}
