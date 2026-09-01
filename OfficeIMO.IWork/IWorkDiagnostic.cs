namespace OfficeIMO.IWork;

/// <summary>One bounded-reader or semantic-projection diagnostic.</summary>
public sealed class IWorkDiagnostic {
    internal IWorkDiagnostic(IWorkDiagnosticSeverity severity, string code, string message,
        string? entryPath = null, ulong? recordIdentifier = null) {
        Severity = severity;
        Code = code;
        Message = message;
        EntryPath = entryPath;
        RecordIdentifier = recordIdentifier;
    }

    /// <summary>Gets the diagnostic severity.</summary>
    public IWorkDiagnosticSeverity Severity { get; }
    /// <summary>Gets the stable diagnostic code.</summary>
    public string Code { get; }
    /// <summary>Gets the human-readable diagnostic message.</summary>
    public string Message { get; }
    /// <summary>Gets the package entry associated with the diagnostic, when available.</summary>
    public string? EntryPath { get; }
    /// <summary>Gets the IWA object identifier associated with the diagnostic, when available.</summary>
    public ulong? RecordIdentifier { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Severity} {Code}: {Message}";
}
