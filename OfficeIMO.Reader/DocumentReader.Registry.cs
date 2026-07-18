using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static string BuildStableId(string kind, string fileName, int chunkIndex, int? blockIndex) {
        // Keep IDs short, stable and ASCII-only; do not leak full paths.
        var l = blockIndex.HasValue ? blockIndex.Value.ToString(CultureInfo.InvariantCulture) : "na";
        return $"{kind}:{fileName}:c{chunkIndex}:l{l}";
    }

    private static MemoryStream CopyToMemory(Stream stream, CancellationToken ct, long? maxInputBytes = null) {
        ct.ThrowIfCancellationRequested();
        var ms = new MemoryStream();
        var buffer = new byte[64 * 1024];
        long totalBytes = 0;
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
            ct.ThrowIfCancellationRequested();
            long nextTotalBytes = totalBytes + read;
            if (maxInputBytes.HasValue && nextTotalBytes > maxInputBytes.Value) {
                ms.Dispose();
                throw new IOException($"Input exceeds MaxInputBytes ({nextTotalBytes.ToString(CultureInfo.InvariantCulture)} > {maxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }

            ms.Write(buffer, 0, read);
            totalBytes = nextTotalBytes;
        }
        ms.Position = 0;
        return ms;
    }

    private static ReaderHandlerCapability CloneCapability(ReaderHandlerCapability capability) {
        return new ReaderHandlerCapability {
            Id = capability.Id,
            DisplayName = capability.DisplayName,
            Description = capability.Description,
            Kind = capability.Kind,
            Extensions = capability.Extensions.ToArray(),
            IsBuiltIn = capability.IsBuiltIn,
            SupportsPath = capability.SupportsPath,
            SupportsStream = capability.SupportsStream,
            SupportsDocumentPath = capability.SupportsDocumentPath,
            SupportsDocumentStream = capability.SupportsDocumentStream,
            SupportsAsyncPath = capability.SupportsAsyncPath,
            SupportsAsyncStream = capability.SupportsAsyncStream,
            SchemaId = capability.SchemaId,
            SchemaVersion = capability.SchemaVersion,
            DefaultMaxInputBytes = capability.DefaultMaxInputBytes,
            WarningBehavior = capability.WarningBehavior,
            DeterministicOutput = capability.DeterministicOutput
        };
    }

    private readonly struct ExtensionOverrideCoverage {
        public ExtensionOverrideCoverage(bool supportsPath, bool supportsStream) {
            SupportsPath = supportsPath;
            SupportsStream = supportsStream;
        }

        public bool SupportsPath { get; }

        public bool SupportsStream { get; }

        public ExtensionOverrideCoverage Add(bool supportsPath, bool supportsStream) {
            return new ExtensionOverrideCoverage(SupportsPath || supportsPath, SupportsStream || supportsStream);
        }
    }

    private static ReaderHandlerCapability? CloneCapabilityWithRemainingSupport(ReaderHandlerCapability capability, IReadOnlyDictionary<string, ExtensionOverrideCoverage>? overriddenExtensions) {
        ReaderHandlerCapability clone = CloneCapability(capability);
        if (overriddenExtensions is null || overriddenExtensions.Count == 0) {
            return clone;
        }

        var extensions = new List<string>(clone.Extensions.Count);
        bool supportsPath = false;
        bool supportsStream = false;
        for (int i = 0; i < clone.Extensions.Count; i++) {
            string extension = clone.Extensions[i];
            if (!overriddenExtensions.TryGetValue(extension, out ExtensionOverrideCoverage coverage)) {
                extensions.Add(extension);
                supportsPath |= clone.SupportsPath;
                supportsStream |= clone.SupportsStream;
                continue;
            }

            bool extensionSupportsPath = clone.SupportsPath && !coverage.SupportsPath;
            bool extensionSupportsStream = clone.SupportsStream && !coverage.SupportsStream;
            if (extensionSupportsPath || extensionSupportsStream) {
                extensions.Add(extension);
                supportsPath |= extensionSupportsPath;
                supportsStream |= extensionSupportsStream;
            }
        }

        if (extensions.Count == 0) {
            return null;
        }

        clone.Extensions = extensions.AsReadOnly();
        clone.SupportsPath = supportsPath;
        clone.SupportsStream = supportsStream;
        return clone;
    }

    internal static string NormalizeExtension(string? extension) {
        var value = extension ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var ext = value.Trim();
        if (!ext.StartsWith(".", StringComparison.Ordinal)) {
            ext = "." + ext;
        }
        return ext.ToLowerInvariant();
    }

    private static HashSet<string> BuildBuiltInExtensionSet() {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var capability in BuiltInCapabilities) {
            foreach (var ext in capability.Extensions) {
                var normalized = NormalizeExtension(ext);
                if (normalized.Length == 0) continue;
                set.Add(normalized);
            }
        }
        return set;
    }

    private static bool TryResolveCustomHandlerByPath(string path, out ReaderHandlerDescriptor handler) {
        var ext = NormalizeExtension(TryGetExtension(path));
        return TryResolveCustomHandlerByExtension(ext, out handler);
    }

    private static bool TryResolveCustomHandlerBySourceName(string? sourceName, out ReaderHandlerDescriptor handler) {
        var ext = NormalizeExtension(TryGetExtension(sourceName ?? string.Empty));
        return TryResolveCustomHandlerByExtension(ext, out handler);
    }

    private static bool TryResolveCustomHandlerByExtension(string ext, out ReaderHandlerDescriptor handler) {
        return GetActiveHandlerRegistry().TryResolve(ext, out handler);
    }

    internal static long? ResolveInitialMaxInputBytes(string? sourceName,
        ReaderOptions options) {
        if (options.MaxInputBytes.HasValue) {
            return options.MaxInputBytes;
        }

        return ResolveHandlerDefaultMaxInputBytes(sourceName);
    }

    internal static long? ResolveStreamMaxInputBytes(string? sourceName,
        ReaderOptions options, bool streamCanSeek) {
        long? resolved = ResolveInitialMaxInputBytes(sourceName, options);
        if (resolved.HasValue || streamCanSeek) return resolved;

        if (TryResolveCustomHandlerBySourceName(sourceName,
                out ReaderHandlerDescriptor customHandler)
            && customHandler.SupportsStreamInput) {
            return null;
        }

        string extension = NormalizeExtension(
            TryGetExtension(sourceName ?? string.Empty));
        for (int i = 0; i < BuiltInCapabilities.Length; i++) {
            ReaderHandlerCapability capability = BuiltInCapabilities[i];
            if (capability.SupportsStream
                && capability.Extensions.Contains(extension,
                    StringComparer.OrdinalIgnoreCase)) {
                return null;
            }
        }

        // Detection needs a seekable snapshot. Keep unidentified streams
        // bounded before their content has established a concrete handler.
        return LegacyPptImportOptions.DefaultMaxInputBytes;
    }

    internal static long? ResolveHandlerDefaultMaxInputBytes(string? sourceName) {
        if (TryResolveCustomHandlerBySourceName(sourceName, out ReaderHandlerDescriptor customHandler)
            && customHandler.SupportsStreamInput) {
            return customHandler.DefaultMaxInputBytes;
        }

        string extension = NormalizeExtension(TryGetExtension(sourceName ?? string.Empty));
        if (extension.Length == 0) {
            return null;
        }

        for (int i = 0; i < BuiltInCapabilities.Length; i++) {
            ReaderHandlerCapability capability = BuiltInCapabilities[i];
            if (capability.SupportsStream
                && capability.DefaultMaxInputBytes.HasValue
                && capability.Extensions.Contains(extension, StringComparer.OrdinalIgnoreCase)) {
                return capability.DefaultMaxInputBytes;
            }
        }

        return null;
    }

    private static bool TryResolveCustomHandlerByKind(ReaderInputKind kind, bool pathInput, out ReaderHandlerDescriptor handler) {
        return GetActiveHandlerRegistry().TryResolveByKind(kind, pathInput, out handler);
    }

    private static ReaderOptions NormalizeOptions(ReaderOptions? options) {
        // Avoid mutating a caller-provided options instance.
        var o = options;
        var clone = new ReaderOptions {
            MaxInputBytes = o?.MaxInputBytes,
            OpenXmlMaxCharactersInPart = o == null ? ReaderOptions.DefaultOpenXmlMaxCharactersInPart : o.OpenXmlMaxCharactersInPart,
            MaxOpenXmlImageAssets = o == null ? ReaderOptions.DefaultMaxOpenXmlImageAssets : o.MaxOpenXmlImageAssets,
            OpenPassword = o?.OpenPassword,
            MaxOpenXmlImagePlacementsPerRelationship = o == null ? ReaderOptions.DefaultMaxOpenXmlImagePlacementsPerRelationship : o.MaxOpenXmlImagePlacementsPerRelationship,
            MaxOpenXmlImageAssetBytes = o == null ? ReaderOptions.DefaultMaxOpenXmlImageAssetBytes : o.MaxOpenXmlImageAssetBytes,
            MaxOpenXmlImageTotalAssetBytes = o == null ? ReaderOptions.DefaultMaxOpenXmlImageTotalAssetBytes : o.MaxOpenXmlImageTotalAssetBytes,
            MaxChars = o?.MaxChars ?? 8_000,
            MaxTableRows = o?.MaxTableRows ?? 200,
            IncludeWordFootnotes = o?.IncludeWordFootnotes ?? true,
            IncludePowerPointNotes = o?.IncludePowerPointNotes ?? true,
            ExcelHeadersInFirstRow = o?.ExcelHeadersInFirstRow ?? true,
            ExcelChunkRows = o?.ExcelChunkRows ?? 200,
            ExcelSheetName = o?.ExcelSheetName,
            ExcelA1Range = o?.ExcelA1Range,
            MarkdownChunkByHeadings = o?.MarkdownChunkByHeadings ?? true,
            MarkdownInputNormalization = CloneMarkdownInputNormalization(o?.MarkdownInputNormalization),
            ComputeHashes = o?.ComputeHashes ?? true,
            DetectionMode = o?.DetectionMode ?? ReaderDetectionMode.ContentWhenUnknown,
            DetectionMaxProbeBytes = o?.DetectionMaxProbeBytes ?? ReaderOptions.DefaultDetectionMaxProbeBytes,
            DetectionMaxContainerEntries = o?.DetectionMaxContainerEntries ?? ReaderOptions.DefaultDetectionMaxContainerEntries
        };

        if (clone.MaxChars < 256) clone.MaxChars = 256;
        if (clone.MaxTableRows < 1) clone.MaxTableRows = 1;
        if (clone.ExcelChunkRows < 1) clone.ExcelChunkRows = 1;
        if (clone.OpenXmlMaxCharactersInPart.HasValue && clone.OpenXmlMaxCharactersInPart.Value < 1) clone.OpenXmlMaxCharactersInPart = null;
        if (clone.MaxOpenXmlImageAssets.HasValue && clone.MaxOpenXmlImageAssets.Value < 1) clone.MaxOpenXmlImageAssets = null;
        if (clone.MaxOpenXmlImagePlacementsPerRelationship.HasValue && clone.MaxOpenXmlImagePlacementsPerRelationship.Value < 1) clone.MaxOpenXmlImagePlacementsPerRelationship = null;
        if (clone.MaxOpenXmlImageAssetBytes.HasValue && clone.MaxOpenXmlImageAssetBytes.Value < 1) clone.MaxOpenXmlImageAssetBytes = null;
        if (clone.MaxOpenXmlImageTotalAssetBytes.HasValue && clone.MaxOpenXmlImageTotalAssetBytes.Value < 1) clone.MaxOpenXmlImageTotalAssetBytes = null;
        if (clone.DetectionMaxProbeBytes < 256) clone.DetectionMaxProbeBytes = 256;
        if (clone.DetectionMaxProbeBytes > ReaderOptions.MaximumDetectionProbeBytes) clone.DetectionMaxProbeBytes = ReaderOptions.MaximumDetectionProbeBytes;
        if (clone.DetectionMaxContainerEntries < 1) clone.DetectionMaxContainerEntries = 1;
        if (clone.DetectionMaxContainerEntries > ReaderOptions.MaximumDetectionContainerEntries) clone.DetectionMaxContainerEntries = ReaderOptions.MaximumDetectionContainerEntries;
        if (!Enum.IsDefined(typeof(ReaderDetectionMode), clone.DetectionMode)) clone.DetectionMode = ReaderDetectionMode.ContentWhenUnknown;

        return clone;
    }

    private static ReaderOptions CloneOptions(ReaderOptions options, bool? computeHashes = null) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        return new ReaderOptions {
            MaxInputBytes = options.MaxInputBytes,
            OpenXmlMaxCharactersInPart = options.OpenXmlMaxCharactersInPart,
            MaxOpenXmlImageAssets = options.MaxOpenXmlImageAssets,
            OpenPassword = options.OpenPassword,
            MaxOpenXmlImagePlacementsPerRelationship = options.MaxOpenXmlImagePlacementsPerRelationship,
            MaxOpenXmlImageAssetBytes = options.MaxOpenXmlImageAssetBytes,
            MaxOpenXmlImageTotalAssetBytes = options.MaxOpenXmlImageTotalAssetBytes,
            MaxChars = options.MaxChars,
            MaxTableRows = options.MaxTableRows,
            IncludeWordFootnotes = options.IncludeWordFootnotes,
            IncludePowerPointNotes = options.IncludePowerPointNotes,
            ExcelHeadersInFirstRow = options.ExcelHeadersInFirstRow,
            ExcelChunkRows = options.ExcelChunkRows,
            ExcelSheetName = options.ExcelSheetName,
            ExcelA1Range = options.ExcelA1Range,
            MarkdownChunkByHeadings = options.MarkdownChunkByHeadings,
            MarkdownInputNormalization = CloneMarkdownInputNormalization(options.MarkdownInputNormalization),
            ComputeHashes = computeHashes ?? options.ComputeHashes,
            DetectionMode = options.DetectionMode,
            DetectionMaxProbeBytes = options.DetectionMaxProbeBytes,
            DetectionMaxContainerEntries = options.DetectionMaxContainerEntries
        };
    }

    private static MarkdownInputNormalizationOptions? CloneMarkdownInputNormalization(MarkdownInputNormalizationOptions? options) {
        if (options == null) {
            return null;
        }

        return new MarkdownInputNormalizationOptions {
            NormalizeSoftWrappedStrongSpans = options.NormalizeSoftWrappedStrongSpans,
            NormalizeInlineCodeSpanLineBreaks = options.NormalizeInlineCodeSpanLineBreaks,
            NormalizeEscapedInlineCodeSpans = options.NormalizeEscapedInlineCodeSpans,
            NormalizeTightStrongBoundaries = options.NormalizeTightStrongBoundaries,
            NormalizeTightArrowStrongBoundaries = options.NormalizeTightArrowStrongBoundaries,
            NormalizeBrokenStrongArrowLabels = options.NormalizeBrokenStrongArrowLabels,
            NormalizeTightColonSpacing = options.NormalizeTightColonSpacing,
            NormalizeHeadingListBoundaries = options.NormalizeHeadingListBoundaries,
            NormalizeCompactStrongLabelListBoundaries = options.NormalizeCompactStrongLabelListBoundaries,
            NormalizeCompactHeadingBoundaries = options.NormalizeCompactHeadingBoundaries,
            NormalizeColonListBoundaries = options.NormalizeColonListBoundaries,
            NormalizeCompactFenceBodyBoundaries = options.NormalizeCompactFenceBodyBoundaries,
            NormalizeLooseStrongDelimiters = options.NormalizeLooseStrongDelimiters,
            NormalizeOrderedListMarkerSpacing = options.NormalizeOrderedListMarkerSpacing,
            NormalizeOrderedListParenMarkers = options.NormalizeOrderedListParenMarkers,
            NormalizeOrderedListCaretArtifacts = options.NormalizeOrderedListCaretArtifacts,
            NormalizeTightParentheticalSpacing = options.NormalizeTightParentheticalSpacing,
            NormalizeNestedStrongDelimiters = options.NormalizeNestedStrongDelimiters
        };
    }

    private static ReaderFolderOptions NormalizeFolderOptions(ReaderFolderOptions? folderOptions) {
        var o = folderOptions;
        var clone = new ReaderFolderOptions {
            Recurse = o?.Recurse ?? true,
            MaxFiles = o?.MaxFiles ?? 500,
            MaxTotalBytes = o?.MaxTotalBytes,
            Extensions = (o?.Extensions == null || o.Extensions.Count == 0) ? null : o.Extensions.ToArray(),
            SkipReparsePoints = o?.SkipReparsePoints ?? true,
            DeterministicOrder = o?.DeterministicOrder ?? true
        };

        if (clone.MaxFiles < 1) clone.MaxFiles = 1;
        if (clone.MaxTotalBytes.HasValue && clone.MaxTotalBytes.Value < 1) clone.MaxTotalBytes = 1;
        return clone;
    }

    private static HashSet<string> NormalizeExtensions(IReadOnlyList<string>? configuredExtensions) {
        var allowedExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var source = (configuredExtensions == null || configuredExtensions.Count == 0)
            ? GetDefaultAndRegisteredFolderExtensions()
            : configuredExtensions;

        foreach (var e in source) {
            if (string.IsNullOrWhiteSpace(e)) continue;
            var normalized = e.StartsWith(".", StringComparison.Ordinal) ? e.Trim() : "." + e.Trim();
            if (normalized.Length > 1) allowedExt.Add(normalized);
        }

        return allowedExt;
    }

    private static IReadOnlyList<string> GetDefaultAndRegisteredFolderExtensions() {
        IReadOnlyList<string> customExtensions = GetActiveHandlerRegistry().Extensions;
        if (customExtensions.Count == 0) {
            return DefaultFolderExtensions;
        }

        return DefaultFolderExtensions.Concat(customExtensions).ToArray();
    }

    private static bool IsDefaultWinmailDat(string path, IReadOnlyList<string>? configuredExtensions) {
        return (configuredExtensions == null || configuredExtensions.Count == 0) &&
            string.Equals(Path.GetFileName(path), "winmail.dat", StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> EnumerateFilesSafeDeterministic(string folderPath, ReaderFolderOptions options, CancellationToken cancellationToken) {
        var dirs = new Queue<string>();
        dirs.Enqueue(folderPath);

        while (dirs.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var dir = dirs.Dequeue();

            IEnumerable<string> entries;
            try {
                entries = Directory.EnumerateFileSystemEntries(dir);
            } catch {
                // Best-effort traversal: unreadable directories are ignored.
                continue;
            }

            var ordered = options.DeterministicOrder
                ? entries.OrderBy(static x => x, StringComparer.Ordinal).ToArray()
                : entries.ToArray();

            foreach (var entry in ordered) {
                cancellationToken.ThrowIfCancellationRequested();

                FileAttributes attrs;
                try {
                    attrs = File.GetAttributes(entry);
                } catch {
                    continue;
                }

                var isDirectory = (attrs & FileAttributes.Directory) == FileAttributes.Directory;
                if (isDirectory) {
                    if (!options.Recurse) continue;

                    if (options.SkipReparsePoints && (attrs & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint) {
                        continue;
                    }

                    dirs.Enqueue(entry);
                    continue;
                }

                yield return entry;
            }
        }
    }

    private static OpenSettings? CreateOpenSettings(ReaderOptions opt) {
        if (opt == null) return null;
        if (!opt.OpenXmlMaxCharactersInPart.HasValue) return null;
        return new OpenSettings {
            AutoSave = false,
            MaxCharactersInPart = opt.OpenXmlMaxCharactersInPart.Value
        };
    }

    private static void EnforceFileSize(string path, long? maxBytes) {
        if (!maxBytes.HasValue) return;
        try {
            var fi = new FileInfo(path);
            if (fi.Length > maxBytes.Value) {
                throw new IOException($"Input exceeds MaxInputBytes ({fi.Length.ToString(CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }
        } catch (IOException) {
            throw;
        } catch {
            // If we can't stat, don't block reads.
        }
    }

    private static string ReadAllText(Stream stream, CancellationToken ct, int? hardCapChars = 50_000_000) {
        ct.ThrowIfCancellationRequested();
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: true);
        var sb = new StringBuilder();
        var buffer = new char[16 * 1024];
        int read;
        while ((read = reader.Read(buffer, 0, buffer.Length)) > 0) {
            ct.ThrowIfCancellationRequested();
            sb.Append(buffer, 0, read);
            if (hardCapChars.HasValue && sb.Length >= hardCapChars.Value) break;
        }
        return sb.ToString();
    }

}
