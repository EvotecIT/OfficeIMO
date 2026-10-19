using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO.Archives;

/// <summary>
/// Safe ZIP traversal helpers for ingestion pipelines.
/// </summary>
public static partial class ArchiveReader {
    /// <summary>
    /// Enumerates ZIP entries from a path.
    /// </summary>
    public static IReadOnlyList<ArchiveEntryDescriptor> EnumerateZip(string zipPath, ArchiveTraversalOptions? options = null) {
        if (zipPath == null) throw new ArgumentNullException(nameof(zipPath));
        if (zipPath.Length == 0) throw new ArgumentException("ZIP path cannot be empty.", nameof(zipPath));
        if (!File.Exists(zipPath)) throw new FileNotFoundException($"ZIP file '{zipPath}' doesn't exist.", zipPath);

        using var fs = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return TraverseZip(fs, options).Entries;
    }

    /// <summary>
    /// Enumerates ZIP entries from a stream.
    /// </summary>
    public static IReadOnlyList<ArchiveEntryDescriptor> EnumerateZip(Stream zipStream, ArchiveTraversalOptions? options = null) {
        return TraverseZip(zipStream, options).Entries;
    }

    /// <summary>
    /// Traverses ZIP entries from a path and returns accepted entries with warnings.
    /// </summary>
    public static ArchiveTraversalResult TraverseZip(string zipPath, ArchiveTraversalOptions? options = null) {
        if (zipPath == null) throw new ArgumentNullException(nameof(zipPath));
        if (zipPath.Length == 0) throw new ArgumentException("ZIP path cannot be empty.", nameof(zipPath));
        if (!File.Exists(zipPath)) throw new FileNotFoundException($"ZIP file '{zipPath}' doesn't exist.", zipPath);

        using var fs = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return TraverseZip(fs, options);
    }

    /// <summary>
    /// Traverses ZIP entries from a stream and returns accepted entries with warnings.
    /// </summary>
    public static ArchiveTraversalResult TraverseZip(Stream zipStream, ArchiveTraversalOptions? options = null) {
        return TraverseZip(zipStream, options, CancellationToken.None);
    }

    /// <summary>Traverses a ZIP source with a physical-entry preflight and cooperative cancellation.</summary>
    public static ArchiveTraversalResult TraverseZip(Stream zipStream, ArchiveTraversalOptions? options,
        CancellationToken cancellationToken) {
        if (zipStream == null) throw new ArgumentNullException(nameof(zipStream));
        if (!zipStream.CanRead) throw new IOException("ZIP stream must be readable.");

        var effective = Normalize(options);
        using Stream snapshot = CreateBoundedSnapshot(zipStream, effective.MaxArchiveBytes, cancellationToken);
        ValidateZipSource(snapshot, effective, cancellationToken);
        using var inspection = ArchiveInspector.OpenZipSnapshot(snapshot, new ArchiveInspectionOptions {
            OperationBudget = effective.OperationBudget, ZipLegacyCodePage = effective.ZipLegacyCodePage,
            MaxInputBytes = effective.MaxArchiveBytes, MaxEntryCount = effective.MaxPhysicalEntries,
            MaxEntryBytes = long.MaxValue, MaxExpandedBytes = long.MaxValue
        }, cancellationToken);
        return TraverseZip(inspection, effective, cancellationToken);
    }

    /// <summary>
    /// Traverses ZIP entries from an already opened archive and returns accepted entries with warnings.
    /// Names are already decoded by its owner; raw filename encoding cannot be validated or changed by this overload.
    /// </summary>
    public static ArchiveTraversalResult TraverseZip(ZipArchive archive, ArchiveTraversalOptions? options = null) {
        return TraverseZip(archive, options, CancellationToken.None);
    }

    /// <summary>Traverses an opened archive with bounded metadata processing and cooperative cancellation.
    /// Names retain the decoding policy applied when the archive was opened.</summary>
    public static ArchiveTraversalResult TraverseZip(ZipArchive archive, ArchiveTraversalOptions? options,
        CancellationToken cancellationToken) {
        if (archive == null) throw new ArgumentNullException(nameof(archive));

        var effective = Normalize(options);
        return TraverseZipCore(archive, effective, cancellationToken);
    }

    private static ArchiveTraversalResult TraverseZipCore(ZipArchive archive, ArchiveTraversalOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (archive.Entries.Count > options.MaxPhysicalEntries)
            throw new InvalidDataException($"ZIP source exceeds MaxPhysicalEntries ({options.MaxPhysicalEntries}).");
        options.OperationBudget?.AdmitEntries(archive.Entries.Count);
        var entries = new List<TraversalEntry>();
        foreach (var entry in archive.Entries) {
            cancellationToken.ThrowIfCancellationRequested();
            entries.Add(new TraversalEntry(entry.FullName, entry.Name, TryGetLength(entry, out long length) ? length : -1,
                entry.CompressedLength, TryGetLastWriteUtc(entry), null));
        }
        return TraverseEntries(entries, options, cancellationToken);
    }

    /// <summary>Applies ZIP ingestion selection policy to an existing immutable inspection without reparsing its input.
    /// Metadata acceptance does not verify member content; use inspection.OpenRead or ReadMembers for payloads.</summary>
    public static ArchiveTraversalResult TraverseZip(ArchiveInspection inspection, ArchiveTraversalOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (inspection == null) throw new ArgumentNullException(nameof(inspection));
        if (inspection.Format != ArchiveFormat.Zip) throw new ArgumentException("A ZIP inspection is required.", nameof(inspection));
        var entries = new List<TraversalEntry>();
        foreach (var entry in inspection.Entries) {
            cancellationToken.ThrowIfCancellationRequested();
            string normalized = entry.RawName.Replace('\\', '/');
            string name = normalized.Substring(normalized.LastIndexOf('/') + 1);
            entries.Add(new TraversalEntry(entry.RawName, name, entry.Length, entry.CompressedLength ?? 0,
                TraversalLastWriteUtc(entry.LastWriteTime), entry.Diagnostic));
        }
        return TraverseEntries(entries, Normalize(options), cancellationToken);
    }

    private static ArchiveTraversalResult TraverseEntries(IReadOnlyList<TraversalEntry> archiveEntries,
        ArchiveTraversalOptions options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (archiveEntries.Count > options.MaxPhysicalEntries) {
            throw new InvalidDataException($"ZIP source exceeds MaxPhysicalEntries ({options.MaxPhysicalEntries}).");
        }
        var list = new List<ArchiveEntryDescriptor>(Math.Min(archiveEntries.Count, options.MaxEntries));
        var warnings = new List<ArchiveTraversalWarning>();
        bool warningsTruncated = false;
        void Warn(ArchiveTraversalWarning warning) {
            if (warnings.Count < options.MaxWarnings) {
                warnings.Add(warning);
            } else if (!warningsTruncated) {
                warnings[options.MaxWarnings - 1] = new ArchiveTraversalWarning {
                    Warning = $"Further ZIP entry warnings were omitted after MaxWarnings ({options.MaxWarnings}) was reached."
                };
                warningsTruncated = true;
            }
        }
        long totalUncompressed = 0;
        int accepted = 0;
        int visited = 0;

        IEnumerable<(TraversalEntry Entry, int Index)> entries = archiveEntries.Select((entry, index) => (entry, index));
        if (options.DeterministicOrder) {
            entries = entries.OrderBy(item => item.Entry.FullName, StringComparer.Ordinal);
        }

        foreach (var (entry, index) in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            visited++;
            if (entry.Diagnostic != null) {
                Warn(new ArchiveTraversalWarning { EntryPath = entry.FullName, Code = entry.Diagnostic.Code, Warning = entry.Diagnostic.Message });
                continue;
            }
            var fullName = NormalizeEntryName(entry.FullName);
            if (fullName.Length == 0) {
                Warn(new ArchiveTraversalWarning {
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
                Warn(new ArchiveTraversalWarning {
                    EntryPath = fullName,
                    Warning = "Skipped ZIP entry because path traversal or absolute path patterns were detected."
                });
                continue;
            }

            var depth = ComputeDepth(fullName, isDirectory);
            if (depth > options.MaxDepth) {
                Warn(new ArchiveTraversalWarning {
                    EntryPath = fullName,
                    Warning = $"Skipped ZIP entry because depth {depth} exceeds MaxDepth ({options.MaxDepth})."
                });
                continue;
            }

            if (accepted >= options.MaxEntries) {
                Warn(new ArchiveTraversalWarning {
                    EntryPath = fullName,
                    Warning = $"Stopped ZIP traversal because MaxEntries ({options.MaxEntries}) was reached."
                });
                break;
            }

            long entryLength = 0;
            if (!isDirectory) {
                if ((entryLength = entry.Length) < 0) {
                    Warn(new ArchiveTraversalWarning {
                        EntryPath = fullName,
                        Warning = "Skipped ZIP entry because uncompressed size could not be read."
                    });
                    continue;
                }

                if (entryLength < 0) {
                    Warn(new ArchiveTraversalWarning {
                        EntryPath = fullName,
                        Warning = "Skipped ZIP entry because its uncompressed size is negative."
                    });
                    continue;
                }

                if (options.MaxEntryUncompressedBytes.HasValue && entryLength > options.MaxEntryUncompressedBytes.Value) {
                    Warn(new ArchiveTraversalWarning {
                        EntryPath = fullName,
                        Warning = $"Skipped ZIP entry because uncompressed size {entryLength} exceeds MaxEntryUncompressedBytes ({options.MaxEntryUncompressedBytes.Value})."
                    });
                    continue;
                }

                if (options.MaxCompressionRatio.HasValue && entryLength > 0 && entry.CompressedLength > 0 && (double)entryLength / entry.CompressedLength > options.MaxCompressionRatio.Value) {
                    Warn(new ArchiveTraversalWarning {
                        EntryPath = fullName,
                        Warning = $"Skipped ZIP entry because compression ratio exceeds MaxCompressionRatio ({options.MaxCompressionRatio.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)})."
                    });
                    continue;
                }

                if (options.MaxTotalUncompressedBytes.HasValue &&
                    entryLength > options.MaxTotalUncompressedBytes.Value - totalUncompressed) {
                    Warn(new ArchiveTraversalWarning {
                        EntryPath = fullName,
                        Warning = $"Stopped ZIP traversal because MaxTotalUncompressedBytes ({options.MaxTotalUncompressedBytes.Value}) would be exceeded."
                    });
                    break;
                }

                totalUncompressed = checked(totalUncompressed + entryLength);
            }

            accepted++;
            var lastWriteUtc = entry.LastWriteUtc;
            list.Add(new ArchiveEntryDescriptor {
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

        return new ArchiveTraversalResult {
            Entries = list,
            Warnings = warnings,
            TotalUncompressedBytes = totalUncompressed,
            EntriesVisited = visited
        };
    }

    private sealed class TraversalEntry {
        internal TraversalEntry(string fullName, string name, long length, long compressed, DateTime time, ArchiveInspectionDiagnostic? diagnostic) {
            FullName = fullName; Name = name; Length = length; CompressedLength = compressed; LastWriteUtc = time; Diagnostic = diagnostic;
        }
        internal string FullName { get; }
        internal string Name { get; }
        internal long Length { get; }
        internal long CompressedLength { get; }
        internal DateTime LastWriteUtc { get; }
        internal ArchiveInspectionDiagnostic? Diagnostic { get; }
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

    private static DateTime TryGetLastWriteUtc(ZipArchiveEntry entry) {
        try {
            return entry.LastWriteTime.UtcDateTime;
        } catch {
            return DateTime.MinValue;
        }
    }

    private static DateTime TraversalLastWriteUtc(DateTimeOffset? dosTime) {
        // Inspection preserves the wall-clock fields with a zero offset. Keep traversal's existing
        // ZipArchive convention: interpret timezone-unspecified DOS fields in the current local zone.
        return dosTime.HasValue ? new DateTimeOffset(dosTime.Value.DateTime).UtcDateTime : DateTime.MinValue;
    }

    private static ArchiveTraversalOptions Normalize(ArchiveTraversalOptions? options) {
        var source = options ?? new ArchiveTraversalOptions();
        var o = new ArchiveTraversalOptions {
            OperationBudget = source.OperationBudget, ZipLegacyCodePage = source.ZipLegacyCodePage,
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
