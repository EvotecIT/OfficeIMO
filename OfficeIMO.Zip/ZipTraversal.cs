using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO.Zip;

/// <summary>
/// Safe ZIP traversal helpers for ingestion pipelines.
/// </summary>
public static partial class ZipTraversal {
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
        return Traverse(zipStream, options, CancellationToken.None);
    }

    /// <summary>Traverses a ZIP source with a physical-entry preflight and cooperative cancellation.</summary>
    public static ZipTraversalResult Traverse(Stream zipStream, ZipTraversalOptions? options,
        CancellationToken cancellationToken) {
        if (zipStream == null) throw new ArgumentNullException(nameof(zipStream));
        if (!zipStream.CanRead) throw new IOException("ZIP stream must be readable.");

        var effective = Normalize(options);
        using Stream snapshot = CreateBoundedSnapshot(zipStream, effective.MaxArchiveBytes, cancellationToken);
        ValidateSource(snapshot, effective, cancellationToken);
        using var archive = new ZipArchive(snapshot, ZipArchiveMode.Read, leaveOpen: true);
        return TraverseCore(archive, effective, cancellationToken);
    }

    /// <summary>
    /// Traverses ZIP entries from an already opened archive and returns accepted entries with warnings.
    /// </summary>
    public static ZipTraversalResult Traverse(ZipArchive archive, ZipTraversalOptions? options = null) {
        return Traverse(archive, options, CancellationToken.None);
    }

    /// <summary>Traverses an opened archive with bounded metadata processing and cooperative cancellation.</summary>
    public static ZipTraversalResult Traverse(ZipArchive archive, ZipTraversalOptions? options,
        CancellationToken cancellationToken) {
        if (archive == null) throw new ArgumentNullException(nameof(archive));

        var effective = Normalize(options);
        return TraverseCore(archive, effective, cancellationToken);
    }

    private static ZipTraversalResult TraverseCore(ZipArchive archive, ZipTraversalOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<ZipArchiveEntry> archiveEntries = archive.Entries;
        if (archiveEntries.Count > options.MaxPhysicalEntries) {
            throw new InvalidDataException($"ZIP source exceeds MaxPhysicalEntries ({options.MaxPhysicalEntries}).");
        }
        var list = new List<ZipEntryDescriptor>(Math.Min(archiveEntries.Count, options.MaxEntries));
        var warnings = new List<ZipTraversalWarning>();
        bool warningsTruncated = false;
        void Warn(ZipTraversalWarning warning) {
            if (warnings.Count < options.MaxWarnings) {
                warnings.Add(warning);
            } else if (!warningsTruncated) {
                warnings[options.MaxWarnings - 1] = new ZipTraversalWarning {
                    Warning = $"Further ZIP entry warnings were omitted after MaxWarnings ({options.MaxWarnings}) was reached."
                };
                warningsTruncated = true;
            }
        }
        long totalUncompressed = 0;
        int accepted = 0;
        int visited = 0;

        IEnumerable<(ZipArchiveEntry Entry, int Index)> entries = archiveEntries.Select((entry, index) => (entry, index));
        if (options.DeterministicOrder) {
            entries = entries.OrderBy(item => item.Entry.FullName, StringComparer.Ordinal);
        }

        foreach (var (entry, index) in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            visited++;
            var fullName = NormalizeEntryName(entry.FullName);
            if (fullName.Length == 0) {
                Warn(new ZipTraversalWarning {
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
                Warn(new ZipTraversalWarning {
                    EntryPath = fullName,
                    Warning = "Skipped ZIP entry because path traversal or absolute path patterns were detected."
                });
                continue;
            }

            var depth = ComputeDepth(fullName, isDirectory);
            if (depth > options.MaxDepth) {
                Warn(new ZipTraversalWarning {
                    EntryPath = fullName,
                    Warning = $"Skipped ZIP entry because depth {depth} exceeds MaxDepth ({options.MaxDepth})."
                });
                continue;
            }

            if (accepted >= options.MaxEntries) {
                Warn(new ZipTraversalWarning {
                    EntryPath = fullName,
                    Warning = $"Stopped ZIP traversal because MaxEntries ({options.MaxEntries}) was reached."
                });
                break;
            }

            long entryLength = 0;
            if (!isDirectory) {
                if (!TryGetLength(entry, out entryLength)) {
                    Warn(new ZipTraversalWarning {
                        EntryPath = fullName,
                        Warning = "Skipped ZIP entry because uncompressed size could not be read."
                    });
                    continue;
                }

                if (entryLength < 0) {
                    Warn(new ZipTraversalWarning {
                        EntryPath = fullName,
                        Warning = "Skipped ZIP entry because its uncompressed size is negative."
                    });
                    continue;
                }

                if (options.MaxEntryUncompressedBytes.HasValue && entryLength > options.MaxEntryUncompressedBytes.Value) {
                    Warn(new ZipTraversalWarning {
                        EntryPath = fullName,
                        Warning = $"Skipped ZIP entry because uncompressed size {entryLength} exceeds MaxEntryUncompressedBytes ({options.MaxEntryUncompressedBytes.Value})."
                    });
                    continue;
                }

                if (options.MaxCompressionRatio.HasValue && IsCompressionRatioExceeded(entry, entryLength, options.MaxCompressionRatio.Value)) {
                    Warn(new ZipTraversalWarning {
                        EntryPath = fullName,
                        Warning = $"Skipped ZIP entry because compression ratio exceeds MaxCompressionRatio ({options.MaxCompressionRatio.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)})."
                    });
                    continue;
                }

                if (options.MaxTotalUncompressedBytes.HasValue &&
                    entryLength > options.MaxTotalUncompressedBytes.Value - totalUncompressed) {
                    Warn(new ZipTraversalWarning {
                        EntryPath = fullName,
                        Warning = $"Stopped ZIP traversal because MaxTotalUncompressedBytes ({options.MaxTotalUncompressedBytes.Value}) would be exceeded."
                    });
                    break;
                }

                totalUncompressed = checked(totalUncompressed + entryLength);
            }

            accepted++;
            var lastWriteUtc = TryGetLastWriteUtc(entry);
            list.Add(new ZipEntryDescriptor {
                FullName = fullName,
                RawFullName = entry.FullName,
                EntryIndex = index,
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
            MaxPhysicalEntries = source.MaxPhysicalEntries,
            MaxArchiveBytes = source.MaxArchiveBytes,
            MaxWarnings = source.MaxWarnings,
            MaxDepth = source.MaxDepth,
            MaxTotalUncompressedBytes = source.MaxTotalUncompressedBytes,
            MaxEntryUncompressedBytes = source.MaxEntryUncompressedBytes,
            MaxCompressionRatio = source.MaxCompressionRatio,
            IncludeDirectoryEntries = source.IncludeDirectoryEntries,
            DeterministicOrder = source.DeterministicOrder
        };

        if (o.MaxEntries < 1) o.MaxEntries = 1;
        if (o.MaxPhysicalEntries < 1) o.MaxPhysicalEntries = 1;
        if (o.MaxArchiveBytes < 1) o.MaxArchiveBytes = 1;
        if (o.MaxWarnings < 1) o.MaxWarnings = 1;
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
