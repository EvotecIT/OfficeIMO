namespace OfficeIMO.Workflows;

/// <summary>A reopenable provider folder whose relative file layout is preserved during assembly.</summary>
public sealed class OfficeWorkflowDirectoryInput {
    /// <summary>Creates a folder reader. Enumeration must yield directories before their descendants and must not follow links.</summary>
    /// <param name="enumerate">Enumerates a fresh view on every call, observing recursion, entry, depth, and cancellation limits. The provider owns item references; the runner closes every opened file stream.</param>
    public OfficeWorkflowDirectoryInput(Func<OfficeWorkflowDirectoryReadOptions, CancellationToken, IAsyncEnumerable<OfficeWorkflowDirectoryEntry>> enumerate) =>
        Enumerate = enumerate ?? throw new ArgumentNullException(nameof(enumerate));

    /// <summary>Gets the enumeration factory. The runner enumerates again before publication to reject changed membership.</summary>
    public Func<OfficeWorkflowDirectoryReadOptions, CancellationToken, IAsyncEnumerable<OfficeWorkflowDirectoryEntry>> Enumerate { get; }
}

/// <summary>Bounds a provider folder traversal, including directories and unsupported files.</summary>
public sealed record OfficeWorkflowDirectoryReadOptions(bool IncludeSubdirectories, int MaximumEntries, int MaximumDepth = 32);

/// <summary>A relative folder member. A null input denotes a directory; otherwise the input opens the selected file.</summary>
/// <param name="RelativePath">Portable slash-separated path below the selected folder, without rooted or dot segments.</param>
/// <param name="Location">Original provider identity used for source/output separation and replacement checks.</param>
/// <param name="Input">Reopenable file access, or null for a directory.</param>
public sealed record OfficeWorkflowDirectoryEntry(string RelativePath, string Location, OfficeWorkflowStreamInput? Input);
