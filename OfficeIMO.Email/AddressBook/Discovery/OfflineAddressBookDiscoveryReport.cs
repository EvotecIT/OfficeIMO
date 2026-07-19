using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

/// <summary>Bounded, content-free inventory of an OAB component file or directory tree.</summary>
public sealed class OfflineAddressBookDiscoveryReport {
    internal OfflineAddressBookDiscoveryReport(
        IReadOnlyList<OfflineAddressBookFileInfo> files,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        Files = files;
        Diagnostics = diagnostics;
    }

    /// <summary>Discovered component metadata in deterministic path order.</summary>
    public IReadOnlyList<OfflineAddressBookFileInfo> Files { get; }
    /// <summary>Skipped directories, reparse points, and other discovery diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>Components that can supply entries through <see cref="OfflineAddressBookSession"/>.</summary>
    public int ReadableFullDetailsCount => Files.Count(file => file.CanEnumerateEntries);
    /// <summary>Recognized legacy, template, and unknown components that are inventoried but not decoded.</summary>
    public int NonEntryComponentCount => Files.Count - ReadableFullDetailsCount;
}
