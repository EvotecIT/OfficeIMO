using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const long DefaultUnidentifiedStreamMaxInputBytes = 64L * 1024L * 1024L;

    internal static string NormalizeExtension(string? extension) {
        string value = extension ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        string normalized = value.Trim();
        if (!normalized.StartsWith(".", StringComparison.Ordinal)) normalized = "." + normalized;
        return normalized.ToLowerInvariant();
    }

    private static string? TryGetExtension(string? path) {
        if (string.IsNullOrWhiteSpace(path)) return null;
        try {
            return Path.GetExtension(path);
        } catch (ArgumentException) {
            return null;
        } catch (NotSupportedException) {
            return null;
        }
    }

    private static bool TryResolveCustomHandlerByPath(string path, out ReaderHandlerDescriptor handler) =>
        TryResolveCustomHandlerByExtension(NormalizeExtension(TryGetExtension(path)), out handler);

    private static bool TryResolveCustomHandlerBySourceName(string? sourceName, out ReaderHandlerDescriptor handler) =>
        TryResolveCustomHandlerByExtension(NormalizeExtension(TryGetExtension(sourceName)), out handler);

    internal static bool CanReadNestedSource(string sourceName) =>
        TryResolveCustomHandlerBySourceName(sourceName, out ReaderHandlerDescriptor handler) &&
        handler.SupportsStreamInput;

    private static bool TryResolveCustomHandlerByExtension(string extension, out ReaderHandlerDescriptor handler) =>
        GetActiveHandlerRegistry().TryResolve(extension, out handler);

    private static bool TryResolveCustomHandlerByKind(ReaderInputKind kind, bool pathInput, out ReaderHandlerDescriptor handler) =>
        GetActiveHandlerRegistry().TryResolveByKind(kind, pathInput, out handler);

    internal static long? ResolveInitialMaxInputBytes(string? sourceName, ReaderOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        return options.MaxInputBytes ?? ResolveHandlerDefaultMaxInputBytes(sourceName);
    }

    internal static long? ResolveStreamMaxInputBytes(string? sourceName, ReaderOptions options, bool streamCanSeek) {
        if (options.MaxInputBytes.HasValue) return options.MaxInputBytes;
        if (TryResolveCustomHandlerBySourceName(sourceName, out ReaderHandlerDescriptor handler) &&
            handler.SupportsStreamInput) return handler.DefaultMaxInputBytes;
        return streamCanSeek ? null : DefaultUnidentifiedStreamMaxInputBytes;
    }

    internal static long? ResolveHandlerDefaultMaxInputBytes(string? sourceName) {
        return TryResolveCustomHandlerBySourceName(sourceName, out ReaderHandlerDescriptor handler)
            ? handler.DefaultMaxInputBytes
            : null;
    }

    private static ReaderOptions NormalizeOptions(ReaderOptions? options) {
        ReaderOptions? source = options;
        var clone = new ReaderOptions {
            MaxInputBytes = source?.MaxInputBytes,
            OpenXmlMaxCharactersInPart = source == null ? ReaderOptions.DefaultOpenXmlMaxCharactersInPart : source.OpenXmlMaxCharactersInPart,
            MaxOpenXmlImageAssets = source == null ? ReaderOptions.DefaultMaxOpenXmlImageAssets : source.MaxOpenXmlImageAssets,
            OpenPassword = source?.OpenPassword,
            MaxOpenXmlImagePlacementsPerRelationship = source == null ? ReaderOptions.DefaultMaxOpenXmlImagePlacementsPerRelationship : source.MaxOpenXmlImagePlacementsPerRelationship,
            MaxOpenXmlImageAssetBytes = source == null ? ReaderOptions.DefaultMaxOpenXmlImageAssetBytes : source.MaxOpenXmlImageAssetBytes,
            MaxOpenXmlImageTotalAssetBytes = source == null ? ReaderOptions.DefaultMaxOpenXmlImageTotalAssetBytes : source.MaxOpenXmlImageTotalAssetBytes,
            MaxChars = source?.MaxChars ?? 8_000,
            MaxTableRows = source?.MaxTableRows ?? 200,
            ComputeHashes = source?.ComputeHashes ?? true,
            DetectionMode = source?.DetectionMode ?? ReaderDetectionMode.ContentWhenUnknown,
            DetectionMaxProbeBytes = source?.DetectionMaxProbeBytes ?? ReaderOptions.DefaultDetectionMaxProbeBytes,
            DetectionMaxContainerEntries = source?.DetectionMaxContainerEntries ?? ReaderOptions.DefaultDetectionMaxContainerEntries
        };

        clone.MaxChars = Math.Max(256, clone.MaxChars);
        clone.MaxTableRows = Math.Max(1, clone.MaxTableRows);
        clone.DetectionMaxProbeBytes = Math.Max(1, Math.Min(ReaderOptions.MaximumDetectionProbeBytes, clone.DetectionMaxProbeBytes));
        clone.DetectionMaxContainerEntries = Math.Max(1, Math.Min(ReaderOptions.MaximumDetectionContainerEntries, clone.DetectionMaxContainerEntries));
        return clone;
    }

    private static ReaderFolderOptions NormalizeFolderOptions(ReaderFolderOptions? options) {
        var clone = new ReaderFolderOptions {
            Recurse = options?.Recurse ?? true,
            MaxFiles = Math.Max(1, options?.MaxFiles ?? 500),
            MaxTotalBytes = options?.MaxTotalBytes,
            Extensions = options?.Extensions == null ? null : options.Extensions.ToArray(),
            SkipReparsePoints = options?.SkipReparsePoints ?? true,
            DeterministicOrder = options?.DeterministicOrder ?? true
        };
        if (clone.MaxTotalBytes.HasValue && clone.MaxTotalBytes.Value < 1) clone.MaxTotalBytes = 1;
        return clone;
    }

    private static HashSet<string> NormalizeExtensions(IReadOnlyList<string>? configuredExtensions) {
        IEnumerable<string> source = configuredExtensions == null || configuredExtensions.Count == 0
            ? GetActiveHandlerRegistry().Extensions
            : configuredExtensions;
        return new HashSet<string>(
            source.Select(NormalizeExtension).Where(static extension => extension.Length > 0),
            StringComparer.OrdinalIgnoreCase);
    }

    private static void EnforceFileSize(string path, long? maxBytes) => ReaderInputLimits.EnforceFileSize(path, maxBytes);

    private static MemoryStream CopyToMemory(Stream stream, CancellationToken cancellationToken, long? maxInputBytes = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        var result = new MemoryStream();
        var buffer = new byte[64 * 1024];
        long total = 0;
        try {
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                total += read;
                if (maxInputBytes.HasValue && total > maxInputBytes.Value) {
                    throw new IOException($"Input exceeds MaxInputBytes ({total.ToString(CultureInfo.InvariantCulture)} > {maxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
                }
                result.Write(buffer, 0, read);
            }
            result.Position = 0;
            return result;
        } catch {
            result.Dispose();
            throw;
        }
    }
}
