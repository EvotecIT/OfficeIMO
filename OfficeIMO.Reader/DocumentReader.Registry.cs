using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Reader;

public static partial class DocumentReader {
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

    private static ReaderHandlerRegistrarDescriptor CloneRegistrarDescriptor(ReaderHandlerRegistrarDescriptor descriptor) {
        return new ReaderHandlerRegistrarDescriptor {
            HandlerId = descriptor.HandlerId,
            AssemblyName = descriptor.AssemblyName,
            TypeName = descriptor.TypeName,
            MethodName = descriptor.MethodName
        };
    }

    private static ReaderHostBootstrapOptions NormalizeHostBootstrapOptions(ReaderHostBootstrapOptions? options) {
        if (options == null) {
            return new ReaderHostBootstrapOptions();
        }

        return new ReaderHostBootstrapOptions {
            ReplaceExistingHandlers = options.ReplaceExistingHandlers,
            IncludeBuiltInCapabilities = options.IncludeBuiltInCapabilities,
            IncludeCustomCapabilities = options.IncludeCustomCapabilities,
            IndentedManifestJson = options.IndentedManifestJson
        };
    }

    private static ReaderHostBootstrapOptions CreateHostBootstrapOptions(ReaderHostBootstrapProfile profile, bool indentedManifestJson) {
        return profile switch {
            ReaderHostBootstrapProfile.ServiceDefault => new ReaderHostBootstrapOptions {
                ReplaceExistingHandlers = true,
                IncludeBuiltInCapabilities = true,
                IncludeCustomCapabilities = true,
                IndentedManifestJson = indentedManifestJson
            },
            ReaderHostBootstrapProfile.ServiceCustomOnly => new ReaderHostBootstrapOptions {
                ReplaceExistingHandlers = true,
                IncludeBuiltInCapabilities = false,
                IncludeCustomCapabilities = true,
                IndentedManifestJson = indentedManifestJson
            },
            ReaderHostBootstrapProfile.ServiceBuiltInOnly => new ReaderHostBootstrapOptions {
                ReplaceExistingHandlers = true,
                IncludeBuiltInCapabilities = true,
                IncludeCustomCapabilities = false,
                IndentedManifestJson = indentedManifestJson
            },
            _ => throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unknown bootstrap profile.")
        };
    }

    private static List<RegistrarCandidate> DiscoverHandlerRegistrarsCore(IEnumerable<Assembly> assemblies) {
        if (assemblies == null) throw new ArgumentNullException(nameof(assemblies));

        var candidates = new List<RegistrarCandidate>();
        var uniqueAssemblies = new Dictionary<string, Assembly>(StringComparer.Ordinal);
        foreach (var assembly in assemblies) {
            if (assembly == null) continue;
            var key = assembly.FullName ?? assembly.GetName().Name ?? assembly.ManifestModule.Name;
            if (!uniqueAssemblies.ContainsKey(key)) {
                uniqueAssemblies.Add(key, assembly);
            }
        }

        var dedupe = new HashSet<string>(StringComparer.Ordinal);
        foreach (var assembly in uniqueAssemblies.Values) {
            foreach (var type in EnumerateLoadableTypes(assembly)) {
                if (type == null) continue;
                if (!type.IsClass || !type.IsAbstract || !type.IsSealed) continue; // static class

                foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)) {
                    if (!IsRegistrarMethod(method, out var handlerId)) continue;

                    var descriptor = new ReaderHandlerRegistrarDescriptor {
                        HandlerId = handlerId,
                        AssemblyName = assembly.GetName().Name ?? string.Empty,
                        TypeName = type.FullName ?? type.Name,
                        MethodName = method.Name
                    };

                    var key = string.Concat(
                        descriptor.AssemblyName, "|",
                        descriptor.TypeName, "|",
                        descriptor.MethodName, "|",
                        descriptor.HandlerId);
                    if (!dedupe.Add(key)) continue;

                    candidates.Add(new RegistrarCandidate(method, descriptor));
                }
            }
        }

        candidates.Sort(static (a, b) => {
            int cmp = string.CompareOrdinal(a.Descriptor.HandlerId, b.Descriptor.HandlerId);
            if (cmp != 0) return cmp;
            cmp = string.CompareOrdinal(a.Descriptor.AssemblyName, b.Descriptor.AssemblyName);
            if (cmp != 0) return cmp;
            cmp = string.CompareOrdinal(a.Descriptor.TypeName, b.Descriptor.TypeName);
            if (cmp != 0) return cmp;
            return string.CompareOrdinal(a.Descriptor.MethodName, b.Descriptor.MethodName);
        });

        return candidates;
    }

    private static IEnumerable<Type> EnumerateLoadableTypes(Assembly assembly) {
        try {
            return assembly.GetTypes();
        } catch (ReflectionTypeLoadException ex) {
            return ex.Types.Where(static t => t != null)!;
        } catch {
            return Array.Empty<Type>();
        }
    }

    private static IReadOnlyList<Assembly> GetLoadedAssembliesByPrefix(string assemblyNamePrefix) {
        if (assemblyNamePrefix == null) throw new ArgumentNullException(nameof(assemblyNamePrefix));

        var prefix = assemblyNamePrefix.Trim();
        if (prefix.Length == 0) {
            throw new ArgumentException("Assembly name prefix cannot be empty.", nameof(assemblyNamePrefix));
        }

        return AppDomain.CurrentDomain.GetAssemblies()
            .Where(static assembly => !assembly.IsDynamic)
            .Where(assembly => (assembly.GetName().Name ?? string.Empty).StartsWith(prefix, StringComparison.Ordinal))
            .OrderBy(static assembly => assembly.GetName().Name ?? string.Empty, StringComparer.Ordinal)
            .ToArray();
    }

    private static bool IsRegistrarMethod(MethodInfo method, out string handlerId) {
        handlerId = string.Empty;
        if (method == null) return false;
        if (method.IsGenericMethodDefinition) return false;
        if (method.ReturnType != typeof(void)) return false;

        var attribute = method.GetCustomAttribute<ReaderHandlerRegistrarAttribute>(inherit: false);
        if (attribute == null) return false;

        handlerId = (attribute.HandlerId ?? string.Empty).Trim();
        if (handlerId.Length == 0) return false;

        bool hasReplaceExisting = false;
        foreach (var parameter in method.GetParameters()) {
            if (parameter.ParameterType == typeof(bool) &&
                string.Equals(parameter.Name, "replaceExisting", StringComparison.OrdinalIgnoreCase)) {
                hasReplaceExisting = true;
                continue;
            }

            if (!parameter.IsOptional) {
                return false;
            }
        }

        return hasReplaceExisting;
    }

    private static string NormalizeExtension(string? extension) {
        var value = extension ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var ext = value.Trim();
        if (!ext.StartsWith(".", StringComparison.Ordinal)) {
            ext = "." + ext;
        }
        return ext.ToLowerInvariant();
    }

    private static List<string> NormalizeRegistrationExtensions(IReadOnlyList<string>? extensions) {
        var list = new List<string>();
        if (extensions == null) return list;

        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var ext in extensions) {
            var normalized = NormalizeExtension(ext);
            if (normalized.Length == 0) continue;
            if (set.Add(normalized)) {
                list.Add(normalized);
            }
        }

        list.Sort(StringComparer.Ordinal);
        return list;
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

    private static bool TryResolveCustomHandlerByPath(string path, out CustomReaderHandler handler) {
        var ext = NormalizeExtension(TryGetExtension(path));
        return TryResolveCustomHandlerByExtension(ext, out handler);
    }

    private static bool TryResolveCustomHandlerBySourceName(string? sourceName, out CustomReaderHandler handler) {
        var ext = NormalizeExtension(TryGetExtension(sourceName ?? string.Empty));
        return TryResolveCustomHandlerByExtension(ext, out handler);
    }

    private static bool TryResolveCustomHandlerByExtension(string ext, out CustomReaderHandler handler) {
        handler = null!;
        if (string.IsNullOrWhiteSpace(ext)) return false;

        lock (HandlerRegistrySync) {
            if (!CustomHandlerIdByExtension.TryGetValue(ext, out var handlerId)) {
                return false;
            }
            if (!CustomHandlersById.TryGetValue(handlerId, out var resolved) || resolved == null) {
                return false;
            }
            handler = resolved;
            return true;
        }
    }

    private static bool RemoveCustomHandlerUnsafe(string handlerId) {
        if (!CustomHandlersById.TryGetValue(handlerId, out var existing)) return false;

        CustomHandlersById.Remove(handlerId);
        foreach (var ext in existing.Extensions) {
            if (CustomHandlerIdByExtension.TryGetValue(ext, out var current) &&
                string.Equals(current, handlerId, StringComparison.OrdinalIgnoreCase)) {
                CustomHandlerIdByExtension.Remove(ext);
            }
        }

        return true;
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
            ComputeHashes = o?.ComputeHashes ?? true
        };

        if (clone.MaxChars < 256) clone.MaxChars = 256;
        if (clone.MaxTableRows < 1) clone.MaxTableRows = 1;
        if (clone.ExcelChunkRows < 1) clone.ExcelChunkRows = 1;
        if (clone.OpenXmlMaxCharactersInPart.HasValue && clone.OpenXmlMaxCharactersInPart.Value < 1) clone.OpenXmlMaxCharactersInPart = null;
        if (clone.MaxOpenXmlImageAssets.HasValue && clone.MaxOpenXmlImageAssets.Value < 1) clone.MaxOpenXmlImageAssets = null;
        if (clone.MaxOpenXmlImagePlacementsPerRelationship.HasValue && clone.MaxOpenXmlImagePlacementsPerRelationship.Value < 1) clone.MaxOpenXmlImagePlacementsPerRelationship = null;
        if (clone.MaxOpenXmlImageAssetBytes.HasValue && clone.MaxOpenXmlImageAssetBytes.Value < 1) clone.MaxOpenXmlImageAssetBytes = null;
        if (clone.MaxOpenXmlImageTotalAssetBytes.HasValue && clone.MaxOpenXmlImageTotalAssetBytes.Value < 1) clone.MaxOpenXmlImageTotalAssetBytes = null;

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
            ComputeHashes = computeHashes ?? options.ComputeHashes
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
        lock (HandlerRegistrySync) {
            if (CustomHandlerIdByExtension.Count == 0) {
                return DefaultFolderExtensions;
            }

            var merged = new string[DefaultFolderExtensions.Length + CustomHandlerIdByExtension.Count];
            Array.Copy(DefaultFolderExtensions, merged, DefaultFolderExtensions.Length);

            int index = DefaultFolderExtensions.Length;
            foreach (var extension in CustomHandlerIdByExtension.Keys) {
                merged[index] = extension;
                index++;
            }

            return merged;
        }
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
