namespace OfficeIMO.Reader.Email;

/// <summary>Item-at-a-time email-store operations for a configured <see cref="OfficeDocumentReader"/>.</summary>
public static class OfficeDocumentReaderEmailStoreExtensions {
    /// <summary>
    /// Lazily reads PST, OST, OLM, EMLX, or a supported mailbox directory one item at a time.
    /// The enumeration uses this Reader instance's modular handlers for semantic body and attachment extraction.
    /// </summary>
    public static IEnumerable<ReaderEmailStoreItemResult> ReadEmailStoreItems(
        this OfficeDocumentReader reader,
        string path,
        ReaderOptions? readerOptions = null,
        ReaderEmailStoreOptions? emailStoreOptions = null,
        CancellationToken cancellationToken = default) =>
        EmailStoreItemReader.Read(reader, path, readerOptions, emailStoreOptions, cancellationToken);

    /// <summary>
    /// Reads and projects one selected email-store item by the stable identifier returned from a store search
    /// or item enumeration. Unrelated item bodies and attachment payloads are not materialized.
    /// </summary>
    public static ReaderEmailStoreItemResult ReadEmailStoreItem(
        this OfficeDocumentReader reader,
        string path,
        string itemId,
        ReaderOptions? readerOptions = null,
        ReaderEmailStoreOptions? emailStoreOptions = null,
        CancellationToken cancellationToken = default) =>
        EmailStoreItemReader.ReadById(
            reader, path, itemId, readerOptions, emailStoreOptions, cancellationToken);
}