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
        if (string.Equals(source, destination, StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidOperationException("Store conversion always writes a different destination file; in-place PST/OST mutation is not supported.");
        }
        using EmailStoreSession session = EmailStoreSession.Open(source,
            readerOptions, cancellationToken);
        return session.ExportToPst(destination, conversionOptions, cancellationToken);
    }
}
