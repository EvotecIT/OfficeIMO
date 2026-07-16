namespace OfficeIMO.Email.AddressBook;

/// <summary>Content-free OAB component discovery and format inspection.</summary>
public static class OfflineAddressBookInspector {
    /// <summary>
    /// Inventories OAB components without requiring a readable Full Details file. Legacy version 2/3 indexes and
    /// display templates are classified but remain explicitly non-enumerable.
    /// </summary>
    public static OfflineAddressBookDiscoveryReport Inspect(
        string path,
        OfflineAddressBookReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        OfflineAddressBookReaderOptions effective = options ?? OfflineAddressBookReaderOptions.Default;
        OabDiscoveryResult result = OabFileDiscovery.Discover(path, effective, cancellationToken);
        return new OfflineAddressBookDiscoveryReport(result.Files, result.Diagnostics);
    }

    /// <summary>Inspects one caller-owned seekable component stream without changing its position.</summary>
    public static OfflineAddressBookFileInfo Inspect(
        Stream stream,
        string sourceName,
        OfflineAddressBookReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfflineAddressBookReaderOptions effective = options ?? OfflineAddressBookReaderOptions.Default;
        OabSource source = OabSource.FromStream(stream, sourceName);
        if (source.Length > effective.MaxInputBytes) {
            throw new OfflineAddressBookLimitExceededException(
                nameof(effective.MaxInputBytes), source.Length, effective.MaxInputBytes, sourceName);
        }
        return OabFileDiscovery.InspectStream(source);
    }
}
