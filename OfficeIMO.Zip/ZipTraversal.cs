using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Zip;

/// <summary>
/// Safe ZIP traversal helpers for ingestion pipelines.
/// </summary>
public static class ZipTraversal {
    /// <summary>
    /// Enumerates ZIP entries from a path.
    /// </summary>
    public static IReadOnlyList<ZipEntryDescriptor> Enumerate(string zipPath, ZipTraversalOptions? options = null) {
        if (zipPath == null) throw new ArgumentNullException(nameof(zipPath));
        if (zipPath.Length == 0) throw new ArgumentException("ZIP path cannot be empty.", nameof(zipPath));
        if (!File.Exists(zipPath)) throw new FileNotFoundException($"ZIP file '{zipPath}' doesn't exist.", zipPath);

        using var fs = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Traverse(fs, options).Entries;
    }

    /// <summary>
    /// Enumerates ZIP entries from a stream.
    /// </summary>
    public static IReadOnlyList<ZipEntryDescriptor> Enumerate(Stream zipStream, ZipTraversalOptions? options = null) {
        return Traverse(zipStream, options).Entries;
    }

    /// <summary>
    /// Traverses ZIP entries from a path and returns accepted entries with warnings.
    /// </summary>
    public static ZipTraversalResult Traverse(string zipPath, ZipTraversalOptions? options = null) {
        if (zipPath == null) throw new ArgumentNullException(nameof(zipPath));
        if (zipPath.Length == 0) throw new ArgumentException("ZIP path cannot be empty.", nameof(zipPath));
        if (!File.Exists(zipPath)) throw new FileNotFoundException($"ZIP file '{zipPath}' doesn't exist.", zipPath);

        using var fs = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Traverse(fs, options);
    }

    /// <summary>
    /// Traverses ZIP entries from a stream and returns accepted entries with warnings.
    /// </summary>
    public static ZipTraversalResult Traverse(Stream zipStream, ZipTraversalOptions? options = null) {
        if (zipStream == null) throw new ArgumentNullException(nameof(zipStream));
        if (!zipStream.CanRead) throw new IOException("ZIP stream must be readable.");

        var effective = Normalize(options);
        using var archive = new ZipArchive(zipStream, ZipArchiveMode.Read, leaveOpen: true);
        return Traverse(archive, effective);
    }

    /// <summary>
    /// Traverses ZIP entries from an already opened archive and returns accepted entries with warnings.
    /// </summary>
    public static ZipTraversalResult Traverse(ZipArchive archive, ZipTraversalOptions? options = null) {
        if (archive == null) throw new ArgumentNullException(nameof(archive));

        var effective = Normalize(options);
        return TraverseCore(archive, effective);
    }

    private static ZipTraversalResult TraverseCore(ZipArchive archive, ZipTraversalOptions options) {
        var list = new List<ZipEntryDescriptor>();
        var warnings = new List<ZipTraversalWarning>();
        long totalUncompressed = 0;
        int accepted = 0;
        int visited = 0;

        IEnumerable<ZipArchiveEntry> entries = archive.Entries;
        if (options.DeterministicOrder) {
            entries = entries.OrderBy(e => e.FullName, StringComparer.Ordinal);
        }

        foreach (var entry in entries) {
            visited++;
            var fullName = NormalizeEntryName(entry.FullName);
            if (fullName.Length == 0) {
                warnings.Add(new ZipTraversalWarning {
                    EntryPath = string.Empty,
                    Warning = "Skipped ZIP entry because its path is empty."
                });
                continue;
            }

            var isDirectory = fullName.EndsWith("/", StringComparison.Ordinal);
            if (isDirectory && !options.IncludeDirectoryEntries) {
                continue;
            }

            if (IsUnsafePath(fullName)) {
                warnings.Add(new ZipTraversalWarning {
                    EntryPath = fullName,
                    Warning = "Skipped ZIP entry because path traversal or absolute path patterns were detected."
                });
                continue;
            }

            var depth = ComputeDepth(fullName, isDirectory);
            if (depth > options.MaxDepth) {
                warnings.Add(new ZipTraversalWarning {
                    EntryPath = fullName,
                    Warning = $"Skipped ZIP entry because depth {depth} exceeds MaxDepth ({options.MaxDepth})."
                });
                continue;
            }

            if (accepted >= options.MaxEntries) {
                warnings.Add(new ZipTraversalWarning {
                    EntryPath = fullName,
                    Warning = $"Stopped ZIP traversal because MaxEntries ({options.MaxEntries}) was reached."
                });
                break;
            }

            long entryLength = 0;
            if (!isDirectory) {
                if (!TryGetLength(entry, out entryLength)) {
                    warnings.Add(new ZipTraversalWarning {
                        EntryPath = fullName,
                        Warning = "Skipped ZIP entry because uncompressed size could not be read."
                    });
                    continue;
                }

                if (options.MaxEntryUncompressedBytes.HasValue && entryLength > options.MaxEntryUncompressedBytes.Value) {
                    warnings.Add(new ZipTraversalWarning {
                        EntryPath = fullName,
                        Warning = $"Skipped ZIP entry because uncompressed size {entryLength} exceeds MaxEntryUncompressedBytes ({options.MaxEntryUncompressedBytes.Value})."
                    });
                    continue;
                }

                if (options.MaxCompressionRatio.HasValue && IsCompressionRatioExceeded(entry, entryLength, options.MaxCompressionRatio.Value)) {
                    warnings.Add(new ZipTraversalWarning {
                        EntryPath = fullName,
                        Warning = $"Skipped ZIP entry because compression ratio exceeds MaxCompressionRatio ({options.MaxCompressionRatio.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)})."
                    });
                    continue;
                }

                if (options.MaxTotalUncompressedBytes.HasValue && (totalUncompressed + entryLength) > options.MaxTotalUncompressedBytes.Value) {
                    warnings.Add(new ZipTraversalWarning {
                        EntryPath = fullName,
                        Warning = $"Stopped ZIP traversal because MaxTotalUncompressedBytes ({options.MaxTotalUncompressedBytes.Value}) would be exceeded."
                    });
                    break;
                }

                totalUncompressed += entryLength;
            }

            accepted++;
            var lastWriteUtc = TryGetLastWriteUtc(entry);
            list.Add(new ZipEntryDescriptor {
                FullName = fullName,
                Name = entry.Name ?? string.Empty,
                IsDirectory = isDirectory,
                Depth = depth,
                UncompressedLength = isDirectory ? 0 : entryLength,
                LastWriteUtc = lastWriteUtc
            });
        }

        return new ZipTraversalResult {
            Entries = list,
            Warnings = warnings,
            TotalUncompressedBytes = totalUncompressed,
            EntriesVisited = visited
        };
    }

    private static string NormalizeEntryName(string? fullName) {
        return OfficeArchiveSafety.NormalizeEntryName(fullName);
    }

    private static bool IsUnsafePath(string fullName) {
        return OfficeArchiveSafety.IsUnsafePath(fullName);
    }

    private static int ComputeDepth(string fullName, bool isDirectory) {
        return OfficeArchiveSafety.ComputeDepth(fullName, isDirectory);
    }

    private static bool TryGetLength(ZipArchiveEntry entry, out long length) {
        return OfficeArchiveSafety.TryGetLength(entry, out length);
    }

    private static bool IsCompressionRatioExceeded(ZipArchiveEntry entry, long uncompressedLength, double maxRatio) {
        return OfficeArchiveSafety.IsCompressionRatioExceeded(entry, uncompressedLength, maxRatio);
    }

    private static DateTime TryGetLastWriteUtc(ZipArchiveEntry entry) {
        try {
            return entry.LastWriteTime.UtcDateTime;
        } catch {
            return DateTime.MinValue;
        }
    }

    private static ZipTraversalOptions Normalize(ZipTraversalOptions? options) {
        var source = options ?? new ZipTraversalOptions();
        var o = new ZipTraversalOptions {
            MaxEntries = source.MaxEntries,
            MaxDepth = source.MaxDepth,
            MaxTotalUncompressedBytes = source.MaxTotalUncompressedBytes,
            MaxEntryUncompressedBytes = source.MaxEntryUncompressedBytes,
            MaxCompressionRatio = source.MaxCompressionRatio,
            IncludeDirectoryEntries = source.IncludeDirectoryEntries,
            DeterministicOrder = source.DeterministicOrder
        };

        if (o.MaxEntries < 1) o.MaxEntries = 1;
        if (o.MaxDepth < 1) o.MaxDepth = 1;
        if (o.MaxTotalUncompressedBytes.HasValue && o.MaxTotalUncompressedBytes.Value < 1) {
            o.MaxTotalUncompressedBytes = 1;
        }
        if (o.MaxEntryUncompressedBytes.HasValue && o.MaxEntryUncompressedBytes.Value < 1) {
            o.MaxEntryUncompressedBytes = 1;
        }
        if (o.MaxCompressionRatio.HasValue && o.MaxCompressionRatio.Value <= 0) {
            o.MaxCompressionRatio = 1;
        }

        return o;
    }
}
