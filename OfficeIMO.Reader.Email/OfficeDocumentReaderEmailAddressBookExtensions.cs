namespace OfficeIMO.Reader.Email;

/// <summary>Address-book operations for a configured <see cref="OfficeDocumentReader"/>.</summary>
public static class OfficeDocumentReaderEmailAddressBookExtensions {
    /// <summary>Lazily reads selected OAB entries one at a time.</summary>
    public static IEnumerable<ReaderEmailAddressBookEntryResult> ReadEmailAddressBookEntries(
        this OfficeDocumentReader reader,
        string path,
        ReaderOptions? readerOptions = null,
        ReaderEmailAddressBookOptions? addressBookOptions = null,
        CancellationToken cancellationToken = default) =>
        EmailAddressBookEntryReader.Read(
            reader, path, readerOptions, addressBookOptions, cancellationToken);
}
