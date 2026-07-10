namespace OfficeIMO.Email;

/// <summary>Result of a bounded email artifact read.</summary>
public sealed class EmailReadResult {
    internal EmailReadResult(EmailDocument document, IReadOnlyList<EmailDiagnostic> diagnostics, long bytesRead) {
        Document = document;
        Diagnostics = diagnostics;
        BytesRead = bytesRead;
    }

    /// <summary>Parsed artifact document.</summary>
    public EmailDocument Document { get; }

    /// <summary>Structured compatibility and fidelity diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>Number of source bytes consumed.</summary>
    public long BytesRead { get; }
}
