using OfficeIMO.ContentSafety;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Epub;

public sealed partial class EpubDocument {
    private const string EpubSignaturePath = "META-INF/signatures.xml";

    /// <summary>Inspects concealed machine-readable HTML across bounded EPUB content documents.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        byte[] packageBytes,
        OfficeContentSafetyOptions? options = null,
        EpubReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        using EpubContentSafetyPackage package = LoadContentSafetyPackage(packageBytes, effective, readOptions, cancellationToken);
        return HtmlContentSafety.InspectPackagePartsAsync(
            "EPUB",
            package.Parts,
            effective,
            cancellationToken).GetAwaiter().GetResult();
    }

    /// <summary>Inspects concealed machine-readable HTML across a bounded EPUB file.</summary>
    public static OfficeContentSafetyReport InspectContentSafety(
        string filePath,
        OfficeContentSafetyOptions? options = null,
        EpubReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        OfficeContentSafetyOptions effective = options ?? new OfficeContentSafetyOptions();
        byte[] input = OfficeContentSafetyInputGuard.ReadAllBytes(filePath, effective, inspectZipPackage: true, cancellationToken);
        return InspectContentSafety(input, effective, readOptions, cancellationToken);
    }

    /// <summary>Removes exact selected concealed HTML findings, preserves other EPUB entries, and reopens the result.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        byte[] packageBytes,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        EpubReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        using EpubContentSafetyPackage package = LoadContentSafetyPackage(packageBytes, options.Inspection, readOptions, cancellationToken);
        HtmlContentSafetyPackageCleanupResult cleaned = HtmlContentSafety.RemoveSelectedPackagePartsAsync(
            "EPUB",
            package.Parts,
            selection,
            options.Inspection,
            cancellationToken).GetAwaiter().GetResult();
        if (cleaned.Changes.Count == 0) {
            return new OfficeContentCleanupResult((byte[])packageBytes.Clone(), cleaned.Before, cleaned.Before, cleaned.Changes);
        }

        bool hasCentralDirectorySignature = OfficeProvenanceZip.HasCentralDirectorySignature(
            packageBytes,
            options.Inspection.MaxPackageEntries,
            cancellationToken);
        bool removeSignature = false;
        if (package.Document.HasSignatures || hasCentralDirectorySignature) {
            if (options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.BlockSave) {
                throw new InvalidOperationException(
                    "EPUB cleanup would invalidate package signature evidence. Select RemoveInvalidatedSignatures explicitly.");
            }
            if (options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.PreserveSignatureMarkup) {
                throw new InvalidOperationException(
                    "EPUB cleanup cannot preserve signature markup as valid evidence after package mutation. Select RemoveInvalidatedSignatures explicitly.");
            }
            removeSignature = package.Document.HasSignatures;
        }

        Dictionary<string, string> replacements = cleaned.Parts.ToDictionary(
            item => item.Key,
            item => item.Value,
            StringComparer.Ordinal);
        var changedParts = new HashSet<string>(cleaned.ChangedParts, StringComparer.Ordinal);
        var xhtmlParts = new HashSet<string>(
            package.Document.Resources.Where(IsXhtmlResource).Select(resource => resource.Path),
            StringComparer.Ordinal);
        OfficeProvenanceSignatureStripResult rewritten = OfficeProvenanceZip.RemoveEntries(
            packageBytes,
            path => removeSignature && path.Equals(EpubSignaturePath, StringComparison.Ordinal),
            options.Inspection.MaxExpandedPackageBytes,
            shouldReplace: path => changedParts.Contains(path),
            replace: (path, original) => EncodeLikeOriginal(replacements[path], original, xhtmlParts.Contains(path)),
            maximumReplacementBytes: options.Inspection.MaxInputBytes,
            maximumOutputBytes: options.Inspection.MaxInputBytes,
            cancellationToken: cancellationToken);
        byte[] output = rewritten.Data;
        OfficeContentSafetyReport after = InspectContentSafety(output, options.Inspection, readOptions, cancellationToken);
        return new OfficeContentCleanupResult(output, cleaned.Before, after, cleaned.Changes);
    }

    /// <summary>Atomically writes an explicitly cleaned EPUB artifact.</summary>
    public static OfficeContentCleanupResult RemoveSelectedContent(
        string inputPath,
        string outputPath,
        OfficeContentCleanupSelection selection,
        OfficeContentCleanupOptions? options = null,
        EpubReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeContentCleanupOptions();
        options.Validate();
        byte[] input = OfficeContentSafetyInputGuard.ReadAllBytes(inputPath, options.Inspection, inspectZipPackage: true, cancellationToken);
        OfficeContentCleanupResult result = RemoveSelectedContent(input, selection, options, readOptions, cancellationToken);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Output);
        return result;
    }

    private static EpubContentSafetyPackage LoadContentSafetyPackage(
        byte[] packageBytes,
        OfficeContentSafetyOptions options,
        EpubReadOptions? readOptions,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeContentSafetyInputGuard.ValidateBytes(packageBytes, options, inspectZipPackage: true);
        OfficeProvenanceZip.ValidateMimetypeEntry(packageBytes, "application/epub+zip", options.MaxPackageEntries);
        ValidateCanonicalEntryPaths(packageBytes, options.MaxPackageEntries, cancellationToken);

        EpubReadOptions effectiveReadOptions = CreateContentSafetyReadOptions(readOptions, options);
        EpubDocument document = EpubReader.ReadBytes(packageBytes, effectiveReadOptions, cancellationToken);
        ThrowForIncompleteContentSafetyRead(document);

        EpubResource? invalidObfuscatedResource = document.Resources.FirstOrDefault(resource =>
            resource.Encryption?.IsFontObfuscation == true
            && !IsFontMediaType(resource.MediaType));
        if (invalidObfuscatedResource != null) {
            throw new InvalidDataException(
                "EPUB content-safety inspection does not support font-obfuscation algorithms on non-font resources: " +
                invalidObfuscatedResource.Path);
        }

        EpubResource[] htmlResources = document.Resources.Where(IsHtmlResource).ToArray();
        if (htmlResources.Length == 0) throw new InvalidDataException("The EPUB manifest contains no local HTML or XHTML content documents.");
        var resourcesByUri = new Dictionary<string, EpubResource>(StringComparer.Ordinal);
        foreach (EpubResource resource in document.Resources.Where(item => !item.IsRemote)) {
            resourcesByUri[CreatePackageUri(resource.Path).AbsoluteUri] = resource;
        }

        var resourceStore = new EpubPackageResourceStore(
            packageBytes,
            resourcesByUri,
            effectiveReadOptions.MaxResourceBytes);
        try {
            var parts = new List<HtmlContentSafetyPackagePart>(htmlResources.Length);
            foreach (EpubResource resource in htmlResources) {
                cancellationToken.ThrowIfCancellationRequested();
                byte[] bytes = resource.Data ?? throw new InvalidDataException(
                    "EPUB content document payload was not retained within the configured limits: " + resource.Path);
                if (resource.Encryption?.RequiresDecryption == true) {
                    throw new InvalidDataException("EPUB content document requires unsupported decryption: " + resource.Path);
                }
                bool xhtml = IsXhtmlResource(resource);
                string html = DecodeContentDocument(bytes, options, xhtml);
                var renderOptions = CreateEpubRenderOptions(
                    resource.Path,
                    resourceStore,
                    effectiveReadOptions,
                    options);
                parts.Add(new HtmlContentSafetyPackagePart(
                    resource.Path,
                    html,
                    "EPUB/" + resource.Path,
                    renderOptions,
                    serializeAsXhtml: xhtml));
            }
            return new EpubContentSafetyPackage(document, parts, resourceStore);
        } catch {
            resourceStore.Dispose();
            throw;
        }
    }

    private static string DecodeContentDocument(
        byte[] bytes,
        OfficeContentSafetyOptions options,
        bool xhtml) {
        if (xhtml) return DecodeXhtmlContentDocument(bytes, options);
        try {
            using var source = new MemoryStream(bytes, writable: false);
            Encoding encoding = HtmlTextEncodingResolver.Default.ResolveHtmlEncoding(source);
            encoding = (Encoding)encoding.Clone();
            encoding.DecoderFallback = DecoderFallback.ExceptionFallback;
            using var reader = new StreamReader(source, encoding, detectEncodingFromByteOrderMarks: true);
            string html = reader.ReadToEnd();
            OfficeContentSafetyInputGuard.ValidateText(html, options);
            return html;
        } catch (DecoderFallbackException exception) {
            throw new InvalidDataException("The EPUB HTML content document contains invalid encoded text.", exception);
        }
    }

    private static string DecodeXhtmlContentDocument(byte[] bytes, OfficeContentSafetyOptions options) {
        OfficeContentSafetyInputGuard.ValidateBytes(bytes, options);
        try {
            Encoding encoding;
            using (var source = new MemoryStream(bytes, writable: false)) {
                encoding = DetectXhtmlEncoding(source);
            }

            encoding = (Encoding)encoding.Clone();
            encoding.DecoderFallback = DecoderFallback.ExceptionFallback;
            byte[] preamble = encoding.GetPreamble();
            int offset = preamble.Length > 0 && bytes.AsSpan().StartsWith(preamble) ? preamble.Length : 0;
            string xhtml = encoding.GetString(bytes, offset, bytes.Length - offset);
            OfficeContentSafetyInputGuard.ValidateText(xhtml, options);
            return xhtml;
        } catch (Exception exception) when (exception is XmlException
                                            || exception is DecoderFallbackException
                                            || exception is ArgumentException) {
            throw new InvalidDataException("The EPUB XHTML content document contains invalid encoded XML.", exception);
        }
    }

    private static Encoding DetectXhtmlEncoding(Stream source) {
        using var detector = new XmlTextReader(source) {
            DtdProcessing = DtdProcessing.Ignore,
            XmlResolver = null
        };
        detector.Read();
        return detector.Encoding ?? new UTF8Encoding(false, true);
    }

    private static EpubReadOptions CreateContentSafetyReadOptions(EpubReadOptions? source, OfficeContentSafetyOptions safety) {
        source ??= new EpubReadOptions();
        return new EpubReadOptions {
            MaxPackageBytes = Math.Min(source.MaxPackageBytes, safety.MaxInputBytes),
            MaxArchiveEntries = Math.Min(source.MaxArchiveEntries, safety.MaxPackageEntries),
            MaxTotalUncompressedBytes = Math.Min(source.MaxTotalUncompressedBytes, safety.MaxExpandedPackageBytes),
            MaxPackageMetadataBytes = Math.Min(source.MaxPackageMetadataBytes, safety.MaxInputBytes),
            MaxMetadataItems = Math.Min(source.MaxMetadataItems, safety.MaxPackageEntries),
            MaxNavigationItems = 1,
            MaxNavigationDepth = 1,
            MaxChapters = 1,
            MaxChapterBytes = Math.Min(source.MaxChapterBytes ?? long.MaxValue, safety.MaxInputBytes),
            MaxTotalRawHtmlBytes = Math.Min(source.MaxTotalRawHtmlBytes, safety.MaxExpandedPackageBytes),
            IncludeRawHtml = false,
            IncludeResourceData = true,
            ResourceDataFilter = static (path, mediaType) =>
                !string.IsNullOrWhiteSpace(mediaType)
                    ? IsHtmlMediaType(mediaType)
                    : IsHtmlPath(path),
            MaxResources = Math.Min(source.MaxResources, safety.MaxPackageEntries),
            MaxResourceBytes = Math.Min(source.MaxResourceBytes, safety.MaxInputBytes),
            MaxTotalResourceBytes = Math.Min(source.MaxTotalResourceBytes, safety.MaxExpandedPackageBytes),
            DeterministicOrder = true,
            PreferSpineOrder = source.PreferSpineOrder,
            IncludeNonLinearSpineItems = true,
            FallbackToHtmlScan = false
        };
    }

    private static HtmlRenderOptions CreateEpubRenderOptions(
        string contentPath,
        EpubPackageResourceStore resourceStore,
        EpubReadOptions readOptions,
        OfficeContentSafetyOptions options) {
        var resourcePolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile();
        resourcePolicy.AllowedUrlSchemes.Add("epub");
        var renderOptions = new HtmlRenderOptions {
            BaseUri = CreatePackageUri(contentPath),
            ResourceUrlPolicy = resourcePolicy,
            MaxInputCharacters = options.MaxCharacters,
            MaxResourceBytes = Math.Min(options.MaxInputBytes, readOptions.MaxResourceBytes),
            MaxTotalResourceBytes = Math.Min(options.MaxExpandedPackageBytes, readOptions.MaxTotalResourceBytes),
            MaxResourceCount = Math.Min(options.MaxPackageEntries, readOptions.MaxResources),
            MaxResourceRequests = Math.Min(options.MaxPackageEntries, readOptions.MaxResources)
        };
        renderOptions.ResourceResolver = (request, token) =>
            Task.FromResult(resourceStore.Resolve(request.Uri, token));
        return renderOptions;
    }

    private static Uri CreatePackageUri(string path) {
        string escaped = string.Join("/", path.Split('/').Select(Uri.EscapeDataString));
        return new Uri("epub://package/" + escaped, UriKind.Absolute);
    }

    private static bool IsHtmlResource(EpubResource resource) =>
        !resource.IsRemote
        && (!string.IsNullOrWhiteSpace(resource.MediaType)
            ? IsHtmlMediaType(resource.MediaType)
            : IsHtmlPath(resource.Path));

    private static bool IsXhtmlResource(EpubResource resource) =>
        !string.IsNullOrWhiteSpace(resource.MediaType)
            ? string.Equals(resource.MediaType, "application/xhtml+xml", StringComparison.OrdinalIgnoreCase)
            : resource.Path.EndsWith(".xhtml", StringComparison.OrdinalIgnoreCase);

    private static bool IsHtmlMediaType(string? mediaType) =>
        string.Equals(mediaType, "application/xhtml+xml", StringComparison.OrdinalIgnoreCase)
        || string.Equals(mediaType, "text/html", StringComparison.OrdinalIgnoreCase);

    private static bool IsHtmlPath(string path) =>
        path.EndsWith(".xhtml", StringComparison.OrdinalIgnoreCase)
        || path.EndsWith(".html", StringComparison.OrdinalIgnoreCase)
        || path.EndsWith(".htm", StringComparison.OrdinalIgnoreCase);

    private static bool IsFontMediaType(string? mediaType) =>
        string.Equals(mediaType, "font/collection", StringComparison.OrdinalIgnoreCase)
            || string.Equals(mediaType, "font/otf", StringComparison.OrdinalIgnoreCase)
            || string.Equals(mediaType, "font/sfnt", StringComparison.OrdinalIgnoreCase)
            || string.Equals(mediaType, "font/ttf", StringComparison.OrdinalIgnoreCase)
            || string.Equals(mediaType, "font/woff", StringComparison.OrdinalIgnoreCase)
            || string.Equals(mediaType, "font/woff2", StringComparison.OrdinalIgnoreCase)
            || string.Equals(mediaType, "application/font-sfnt", StringComparison.OrdinalIgnoreCase)
            || string.Equals(mediaType, "application/font-woff", StringComparison.OrdinalIgnoreCase)
            || string.Equals(mediaType, "application/vnd.ms-opentype", StringComparison.OrdinalIgnoreCase);

    private static void ThrowForIncompleteContentSafetyRead(EpubDocument document) {
        EpubDiagnostic? missingHtml = document.Diagnostics.FirstOrDefault(item =>
            item.Code.Equals("epub.resource.missing", StringComparison.Ordinal)
            && (!string.IsNullOrWhiteSpace(item.MediaType)
                ? IsHtmlMediaType(item.MediaType)
                : IsHtmlPath(item.Path ?? string.Empty)));
        if (missingHtml != null) {
            throw new InvalidDataException(
                "EPUB content-safety inspection requires every local manifest HTML resource. " + missingHtml.Message);
        }

        string[] blockingPrefixes = {
            "epub.archive.unsafe-path",
            "epub.archive.duplicate-path",
            "epub.container.rootfile-fallback",
            "epub.container.structure-invalid",
            "epub.container.rootfile-path-invalid",
            "epub.container.rootfile-duplicate",
            "epub.container.rootfile-missing",
            "epub.container.multiple-rootfiles",
            "epub.package.missing",
            "epub.package.metadata-size-limit",
            "epub.package.invalid-xml",
            "epub.manifest.invalid-path",
            "epub.manifest.reference-non-conforming",
            "epub.manifest.duplicate-id",
            "epub.manifest.duplicate-target",
            "epub.resource.count-limit",
            "epub.resource.size-limit",
            "epub.resource.total-size-limit",
            "epub.resource.encrypted",
            "epub.encryption.invalid-xml",
            "epub.encryption.metadata-size-limit",
            "epub.encryption.resource-path-invalid",
            "epub.encryption.duplicate-resource",
            "epub.encryption.resource-missing"
        };
        EpubDiagnostic? blocking = document.Diagnostics.FirstOrDefault(item =>
            blockingPrefixes.Any(code => item.Code.Equals(code, StringComparison.Ordinal)));
        if (blocking != null) {
            throw new InvalidDataException("EPUB content-safety inspection requires a complete package projection. " + blocking.Message);
        }
        if (string.IsNullOrWhiteSpace(document.OpfPath)) {
            throw new InvalidDataException("EPUB content-safety inspection requires a readable declared OPF package document.");
        }
    }

    private static void ValidateCanonicalEntryPaths(byte[] packageBytes, int maximumEntries, CancellationToken cancellationToken) {
        using var stream = new MemoryStream(packageBytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        if (archive.Entries.Count > maximumEntries) throw new InvalidDataException("The EPUB package exceeds the configured entry-count limit.");
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (ZipArchiveEntry entry in archive.Entries) {
            cancellationToken.ThrowIfCancellationRequested();
            bool isDirectory = entry.FullName.EndsWith("/", StringComparison.Ordinal);
            if (!EpubReader.TryNormalizeArchiveEntryPath(entry.FullName, out string normalized)
                || !(isDirectory
                    ? string.Concat(normalized, "/").Equals(entry.FullName, StringComparison.Ordinal)
                    : normalized.Equals(entry.FullName, StringComparison.Ordinal))) {
                throw new InvalidDataException("EPUB package contains an unsafe or non-canonical entry path: " + entry.FullName);
            }
            string identity = isDirectory ? string.Concat(normalized, "/") : normalized;
            if (!seen.Add(identity)) throw new InvalidDataException("EPUB package contains a duplicate entry path: " + identity);
        }
    }

    private static byte[] EncodeLikeOriginal(string text, byte[] original, bool xhtml) {
        Encoding encoding;
        byte[] preamble;
        if (original.Length >= 4 && original[0] == 0x00 && original[1] == 0x00 && original[2] == 0xFE && original[3] == 0xFF) {
            encoding = new UTF32Encoding(true, false, true);
            preamble = new byte[] { 0x00, 0x00, 0xFE, 0xFF };
        } else if (original.Length >= 4 && original[0] == 0xFF && original[1] == 0xFE && original[2] == 0x00 && original[3] == 0x00) {
            encoding = new UTF32Encoding(false, false, true);
            preamble = new byte[] { 0xFF, 0xFE, 0x00, 0x00 };
        } else if (original.Length >= 3 && original[0] == 0xEF && original[1] == 0xBB && original[2] == 0xBF) {
            encoding = new UTF8Encoding(false, true);
            preamble = new byte[] { 0xEF, 0xBB, 0xBF };
        } else if (original.Length >= 2 && original[0] == 0xFE && original[1] == 0xFF) {
            encoding = new UnicodeEncoding(true, false, true);
            preamble = new byte[] { 0xFE, 0xFF };
        } else if (original.Length >= 2 && original[0] == 0xFF && original[1] == 0xFE) {
            encoding = new UnicodeEncoding(false, false, true);
            preamble = new byte[] { 0xFF, 0xFE };
        } else if (xhtml) {
            try {
                using var stream = new MemoryStream(original, writable: false);
                encoding = DetectXhtmlEncoding(stream);
                preamble = Array.Empty<byte>();
            } catch (XmlException exception) {
                throw new InvalidDataException("The EPUB XHTML content document encoding cannot be preserved safely.", exception);
            }
        } else {
            using var stream = new MemoryStream(original, writable: false);
            encoding = HtmlTextEncodingResolver.Default.ResolveHtmlEncoding(stream);
            preamble = Array.Empty<byte>();
        }
        try {
            encoding = (Encoding)encoding.Clone();
            encoding.EncoderFallback = EncoderFallback.ExceptionFallback;
        } catch (NotSupportedException exception) {
            throw new InvalidDataException("The EPUB content document declared an encoding that cannot be preserved safely.", exception);
        }
        byte[] payload;
        try {
            payload = encoding.GetBytes(EscapeUnrepresentableHtmlCharacters(text, encoding));
        } catch (EncoderFallbackException exception) {
            throw new InvalidDataException("The EPUB content document encoding cannot represent the cleaned HTML.", exception);
        }
        if (preamble.Length == 0) return payload;
        var result = new byte[preamble.Length + payload.Length];
        Buffer.BlockCopy(preamble, 0, result, 0, preamble.Length);
        Buffer.BlockCopy(payload, 0, result, preamble.Length, payload.Length);
        return result;
    }

    private static string EscapeUnrepresentableHtmlCharacters(string text, Encoding encoding) {
        StringBuilder? escaped = null;
        for (int index = 0; index < text.Length;) {
            int characterCount = 1;
            int codePoint;
            char current = text[index];
            if (char.IsHighSurrogate(current)) {
                if (index + 1 >= text.Length || !char.IsLowSurrogate(text[index + 1])) {
                    throw new InvalidDataException("The cleaned EPUB HTML contains an invalid Unicode surrogate.");
                }
                characterCount = 2;
                codePoint = char.ConvertToUtf32(current, text[index + 1]);
            } else if (char.IsLowSurrogate(current)) {
                throw new InvalidDataException("The cleaned EPUB HTML contains an invalid Unicode surrogate.");
            } else {
                codePoint = current;
            }

            bool representable;
            try {
                encoding.GetByteCount(text.Substring(index, characterCount));
                representable = true;
            } catch (EncoderFallbackException) {
                representable = false;
            }

            if (!representable) {
                if (escaped == null) {
                    escaped = new StringBuilder(text.Length + 16);
                    escaped.Append(text, 0, index);
                }
                escaped.Append("&#x");
                escaped.Append(codePoint.ToString("X", System.Globalization.CultureInfo.InvariantCulture));
                escaped.Append(';');
            } else if (escaped != null) {
                escaped.Append(text, index, characterCount);
            }
            index += characterCount;
        }
        return escaped?.ToString() ?? text;
    }

    private sealed class EpubPackageResourceStore : IDisposable {
        private readonly MemoryStream _stream;
        private readonly ZipArchive _archive;
        private readonly IReadOnlyDictionary<string, EpubResource> _resourcesByUri;
        private readonly Dictionary<string, ZipArchiveEntry> _entriesByPath;
        private readonly Dictionary<string, byte[]> _cache = new(StringComparer.Ordinal);
        private readonly long _maximumBytes;
        private readonly object _sync = new();
        private bool _disposed;

        internal EpubPackageResourceStore(
            byte[] packageBytes,
            IReadOnlyDictionary<string, EpubResource> resourcesByUri,
            long maximumBytes) {
            _stream = new MemoryStream(packageBytes, writable: false);
            _archive = new ZipArchive(_stream, ZipArchiveMode.Read, leaveOpen: true);
            _resourcesByUri = resourcesByUri;
            _entriesByPath = _archive.Entries.ToDictionary(entry => entry.FullName, StringComparer.Ordinal);
            _maximumBytes = maximumBytes;
        }

        internal HtmlResolvedResource? Resolve(Uri requestUri, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            var lookupUri = new UriBuilder(requestUri) {
                Fragment = string.Empty,
                Query = string.Empty
            }.Uri;
            if (!requestUri.Scheme.Equals("epub", StringComparison.OrdinalIgnoreCase)
                || !_resourcesByUri.TryGetValue(lookupUri.AbsoluteUri, out EpubResource? resource)
                || resource.Encryption?.RequiresDecryption == true) {
                return null;
            }

            byte[] data;
            lock (_sync) {
                cancellationToken.ThrowIfCancellationRequested();
                if (_disposed) throw new ObjectDisposedException(nameof(EpubPackageResourceStore));
                if (!_cache.TryGetValue(resource.Path, out data!)) {
                    data = ReadResource(resource, cancellationToken);
                    _cache.Add(resource.Path, data);
                }
            }
            return new HtmlResolvedResource(
                data,
                string.IsNullOrWhiteSpace(resource.MediaType) ? "application/octet-stream" : resource.MediaType!);
        }

        private byte[] ReadResource(EpubResource resource, CancellationToken cancellationToken) {
            if (resource.LengthBytes > _maximumBytes || resource.LengthBytes > int.MaxValue) {
                throw new InvalidDataException("EPUB resource exceeds the configured resource limit: " + resource.Path);
            }
            if (!_entriesByPath.TryGetValue(resource.Path, out ZipArchiveEntry? entry)) {
                throw new InvalidDataException("EPUB resource entry is missing: " + resource.Path);
            }
            if (entry.Length > _maximumBytes || entry.Length > int.MaxValue) {
                throw new InvalidDataException("EPUB resource exceeds the configured resource limit: " + resource.Path);
            }

            using Stream source = entry.Open();
            using var output = new MemoryStream(checked((int)entry.Length));
            var buffer = new byte[81920];
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                if (output.Length > _maximumBytes - read) {
                    throw new InvalidDataException("EPUB resource exceeds the configured resource limit: " + resource.Path);
                }
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }

        public void Dispose() {
            lock (_sync) {
                if (_disposed) return;
                _disposed = true;
                _archive.Dispose();
                _stream.Dispose();
                _cache.Clear();
            }
        }
    }

    private sealed class EpubContentSafetyPackage : IDisposable {
        internal EpubContentSafetyPackage(
            EpubDocument document,
            IReadOnlyList<HtmlContentSafetyPackagePart> parts,
            EpubPackageResourceStore resourceStore) {
            Document = document;
            Parts = parts;
            ResourceStore = resourceStore;
        }

        internal EpubDocument Document { get; }
        internal IReadOnlyList<HtmlContentSafetyPackagePart> Parts { get; }
        private EpubPackageResourceStore ResourceStore { get; }

        public void Dispose() => ResourceStore.Dispose();
    }
}
