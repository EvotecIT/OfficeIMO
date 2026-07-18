namespace OfficeIMO.Email.Store;

/// <summary>Dependency-free mailbox-store conversion entry points.</summary>
public static class EmailStoreConverter {
    /// <summary>Opens a PST, OST, OLM, MBOX, EML, or mailbox directory and writes a new Unicode PST.</summary>
    public static EmailStorePstConversionReport ConvertToPst(string sourcePath,
        string destinationPath, EmailStoreReaderOptions? readerOptions = null,
        EmailStorePstConversionOptions? conversionOptions = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(sourcePath)) throw new ArgumentException("A source path is required.", nameof(sourcePath));
        if (string.IsNullOrWhiteSpace(destinationPath)) throw new ArgumentException("A destination path is required.", nameof(destinationPath));
        string source = Path.GetFullPath(sourcePath);
        string destination = Path.GetFullPath(destinationPath);
        if (EmailStorePathIdentity.AreEquivalent(source, destination)) {
            throw new InvalidOperationException("Store conversion always writes a different destination file; in-place PST/OST mutation is not supported.");
        }
        using EmailStoreSession session = EmailStoreSession.Open(source,
            readerOptions, cancellationToken);
        return session.ExportToPst(destination, conversionOptions, cancellationToken);
    }

    /// <summary>
    /// Merges multiple read-only PST, OST, OLM, EMLX, or mailbox-directory sources into a new Unicode PST.
    /// Folder mapping, semantic deduplication, transient source retries, and diagnostics are handled by the
    /// reusable store engine; sources are never modified.
    /// </summary>
    public static EmailStorePstMergeReport MergeToPst(IEnumerable<EmailStoreMergeSource> sources,
        string destinationPath, EmailStorePstMergeOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(destinationPath)) {
            throw new ArgumentException("A destination path is required.", nameof(destinationPath));
        }
        return new EmailStorePstMerger(sources, destinationPath,
            options ?? new EmailStorePstMergeOptions(), cancellationToken).Run();
    }
}
