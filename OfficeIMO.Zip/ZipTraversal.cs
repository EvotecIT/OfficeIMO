using System.Threading;
using OfficeIMO.Archives;

namespace OfficeIMO.Zip;

/// <summary>Compatibility surface over the canonical OfficeIMO.Archives ZIP traversal engine.</summary>
public static partial class ZipTraversal {
    /// <summary>Enumerates accepted entries from a path.</summary>
    public static IReadOnlyList<ZipEntryDescriptor> Enumerate(string zipPath, ZipTraversalOptions? options = null) => Traverse(zipPath, options).Entries;
    /// <summary>Enumerates accepted entries from a stream.</summary>
    public static IReadOnlyList<ZipEntryDescriptor> Enumerate(Stream zipStream, ZipTraversalOptions? options = null) => Traverse(zipStream, options).Entries;
    /// <summary>Traverses a path using bounded ingestion rules and warnings.</summary>
    public static ZipTraversalResult Traverse(string zipPath, ZipTraversalOptions? options = null) => From(ArchiveReader.TraverseZip(zipPath, Options(options)));
    /// <summary>Traverses a stream using a bounded immutable snapshot.</summary>
    public static ZipTraversalResult Traverse(Stream zipStream, ZipTraversalOptions? options = null) => Traverse(zipStream, options, CancellationToken.None);
    /// <summary>Traverses a stream with cooperative cancellation.</summary>
    public static ZipTraversalResult Traverse(Stream zipStream, ZipTraversalOptions? options, CancellationToken cancellationToken) => From(ArchiveReader.TraverseZip(zipStream, Options(options), cancellationToken));
    /// <summary>Traverses an opened archive's bounded metadata.</summary>
    public static ZipTraversalResult Traverse(ZipArchive archive, ZipTraversalOptions? options = null) => Traverse(archive, options, CancellationToken.None);
    /// <summary>Traverses an opened archive's bounded metadata with cooperative cancellation.</summary>
    public static ZipTraversalResult Traverse(ZipArchive archive, ZipTraversalOptions? options, CancellationToken cancellationToken) => From(ArchiveReader.TraverseZip(archive, Options(options), cancellationToken));

    private static ArchiveTraversalOptions? Options(ZipTraversalOptions? value) => value is null ? null : new() {
        MaxEntries = value.MaxEntries, MaxPhysicalEntries = value.MaxPhysicalEntries, MaxArchiveBytes = value.MaxArchiveBytes,
        MaxWarnings = value.MaxWarnings, MaxDepth = value.MaxDepth, MaxTotalUncompressedBytes = value.MaxTotalUncompressedBytes,
        MaxEntryUncompressedBytes = value.MaxEntryUncompressedBytes, MaxCompressionRatio = value.MaxCompressionRatio,
        IncludeDirectoryEntries = value.IncludeDirectoryEntries, DeterministicOrder = value.DeterministicOrder
    };
    private static ZipTraversalResult From(ArchiveTraversalResult value) => new() {
        Entries = value.Entries.Select(entry => new ZipEntryDescriptor { FullName = entry.FullName, RawFullName = entry.RawFullName,
            EntryIndex = entry.EntryIndex, Name = entry.Name, IsDirectory = entry.IsDirectory, Depth = entry.Depth,
            UncompressedLength = entry.UncompressedLength, LastWriteUtc = entry.LastWriteUtc }).ToArray(),
        Warnings = value.Warnings.Select(warning => new ZipTraversalWarning { EntryPath = warning.EntryPath, Warning = warning.Warning }).ToArray(),
        EntriesVisited = value.EntriesVisited, TotalUncompressedBytes = value.TotalUncompressedBytes
    };
}
