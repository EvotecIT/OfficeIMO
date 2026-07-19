using OfficeIMO.Email;
using OfficeIMO.Email.AddressBook;

namespace OfficeIMO.Reader.Email;

/// <summary>Reader projection of one Offline Address Book entry.</summary>
public sealed class ReaderEmailAddressBookEntryResult {
    internal ReaderEmailAddressBookEntryResult(
        OfflineAddressBookEntryReference reference,
        OfflineAddressBookEntrySummary? summary,
        string logicalPath,
        IReadOnlyList<ReaderChunk> chunks,
        IReadOnlyList<EmailDiagnostic> entryDiagnostics,
        IReadOnlyList<EmailDiagnostic>? sessionDiagnostics = null) {
        Reference = reference;
        Summary = summary;
        LogicalPath = logicalPath;
        Chunks = chunks;
        EntryDiagnostics = entryDiagnostics;
        SessionDiagnostics = sessionDiagnostics ?? Array.Empty<EmailDiagnostic>();
        Diagnostics = SessionDiagnostics.Concat(EntryDiagnostics).ToArray();
    }

    /// <summary>Stable reference within the open session snapshot.</summary>
    public OfflineAddressBookEntryReference Reference { get; }
    /// <summary>Lightweight typed summary when the entry could be decoded.</summary>
    public OfflineAddressBookEntrySummary? Summary { get; }
    /// <summary>Logical list/entry path used for chunk citations without embedding directory values.</summary>
    public string LogicalPath { get; }
    /// <summary>Chunks for this entry only.</summary>
    public IReadOnlyList<ReaderChunk> Chunks { get; }
    /// <summary>Entry-scoped parsing and projection diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> EntryDiagnostics { get; }
    /// <summary>Session-open diagnostics, attached to the first emitted result only.</summary>
    public IReadOnlyList<EmailDiagnostic> SessionDiagnostics { get; }
    /// <summary>Combined session and entry diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>Whether the entry produced a chunk without an entry-scoped error.</summary>
    public bool Succeeded => Chunks.Count > 0 &&
        !EntryDiagnostics.Any(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
}
