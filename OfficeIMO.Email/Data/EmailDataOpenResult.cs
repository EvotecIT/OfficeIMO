using OfficeIMO.Email.AddressBook;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Data;

/// <summary>
/// One typed artifact returned by its existing OfficeIMO owner. Dispose the result to close sessions and retained
/// streaming email content.
/// </summary>
public sealed class EmailDataOpenResult : IDisposable {
    private readonly IDisposable? _ownedResource;
    private bool _disposed;

    internal EmailDataOpenResult(string sourcePath, EmailReadResult email) {
        SourcePath = sourcePath;
        Kind = EmailDataArtifactKind.EmailDocument;
        Email = email;
        _ownedResource = email;
    }

    internal EmailDataOpenResult(string sourcePath, IcsDocument calendar) {
        SourcePath = sourcePath;
        Kind = EmailDataArtifactKind.Calendar;
        Calendar = calendar;
    }

    internal EmailDataOpenResult(string sourcePath, VCardDocument contact) {
        SourcePath = sourcePath;
        Kind = EmailDataArtifactKind.Contact;
        Contact = contact;
    }

    internal EmailDataOpenResult(string sourcePath, EmailStoreSession store) {
        SourcePath = sourcePath;
        Kind = EmailDataArtifactKind.Store;
        Store = store;
        _ownedResource = store;
    }

    internal EmailDataOpenResult(string sourcePath, OfflineAddressBookSession addressBook) {
        SourcePath = sourcePath;
        Kind = EmailDataArtifactKind.OfflineAddressBook;
        AddressBook = addressBook;
        _ownedResource = addressBook;
    }

    /// <summary>Normalized path supplied to the selected owner.</summary>
    public string SourcePath { get; }
    /// <summary>Selected artifact owner and result type.</summary>
    public EmailDataArtifactKind Kind { get; }
    /// <summary>Individual-email read result, including fidelity diagnostics.</summary>
    public EmailReadResult? Email { get; }
    /// <summary>Individual email document convenience projection.</summary>
    public EmailDocument? EmailDocument => Email?.Document;
    /// <summary>Parsed iCalendar document.</summary>
    public IcsDocument? Calendar { get; }
    /// <summary>Parsed vCard document.</summary>
    public VCardDocument? Contact { get; }
    /// <summary>Open mailbox-store session.</summary>
    public EmailStoreSession? Store { get; }
    /// <summary>Open Offline Address Book session.</summary>
    public OfflineAddressBookSession? AddressBook { get; }

    /// <summary>The selected owner result as its public OfficeIMO type.</summary>
    public object Artifact => (object?)EmailDocument ?? Calendar ?? Contact ??
        (object?)Store ?? AddressBook ?? throw new InvalidOperationException("No artifact is available.");

    /// <summary>Closes a store/OAB session or releases file-backed email content. Parsed ICS/VCF objects need no disposal.</summary>
    public void Dispose() {
        if (_disposed) return;
        _ownedResource?.Dispose();
        _disposed = true;
    }
}
