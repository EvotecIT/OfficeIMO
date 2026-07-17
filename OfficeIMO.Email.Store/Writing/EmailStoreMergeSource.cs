namespace OfficeIMO.Email.Store;

/// <summary>One read-only PST, OST, OLM, EMLX, or mailbox-directory source in a PST merge.</summary>
public sealed class EmailStoreMergeSource {
    /// <summary>Creates a merge source.</summary>
    public EmailStoreMergeSource(string path, string? displayName = null,
        EmailStoreReaderOptions? readerOptions = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A source path is required.", nameof(path));
        Path = System.IO.Path.GetFullPath(path);
        DisplayName = string.IsNullOrWhiteSpace(displayName) ? null : displayName!.Trim();
        ReaderOptions = readerOptions;
    }

    /// <summary>Absolute source file or directory path.</summary>
    public string Path { get; }
    /// <summary>Optional destination root-folder name for this source.</summary>
    public string? DisplayName { get; }
    /// <summary>Optional source-specific bounded reader policy.</summary>
    public EmailStoreReaderOptions? ReaderOptions { get; }
}
