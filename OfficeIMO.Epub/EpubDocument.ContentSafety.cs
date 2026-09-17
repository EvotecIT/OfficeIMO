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
        EpubContentSafetyPackage package = LoadContentSafetyPackage(packageBytes, effective, readOptions, cancellationToken);
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
        EpubContentSafetyPackage package = LoadContentSafetyPackage(packageBytes, options.Inspection, readOptions, cancellationToken);
        HtmlContentSafetyPackageCleanupResult cleaned = HtmlContentSafety.RemoveSelectedPackagePartsAsync(
            "EPUB",
            package.Parts,
            selection,
            options.Inspection,
            cancellationToken).GetAwaiter().GetResult();
        if (cleaned.Changes.Count == 0) {
            return new OfficeContentCleanupResult((byte[])packageBytes.Clone(), cleaned.Before, cleaned.Before, cleaned.Changes);
        }

        bool removeSignature = false;
        if (package.Document.HasSignatures) {
            if (options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.BlockSave) {
                throw new InvalidOperationException(
                    "EPUB cleanup would invalidate META-INF/signatures.xml. Select RemoveInvalidatedSignatures explicitly.");
            }
            if (options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.PreserveSignatureMarkup) {
                throw new InvalidOperationException(
                    "EPUB cleanup cannot preserve signature markup as valid evidence after package mutation. Select RemoveInvalidatedSignatures explicitly.");
            }
            removeSignature = true;
        }

        Dictionary<string, string> replacements = cleaned.Parts.ToDictionary(
            item => item.Key,
            item => item.Value,
            StringComparer.Ordinal);
        var changedParts = new HashSet<string>(cleaned.ChangedParts, StringComparer.Ordinal);
        OfficeProvenanceSignatureStripResult rewritten = OfficeProvenanceZip.RemoveEntries(
            packageBytes,
            path => removeSignature && path.Equals(EpubSignaturePath, StringComparison.Ordinal),
            options.Inspection.MaxExpandedPackageBytes,
            shouldReplace: path => changedParts.Contains(path),
            replace: (path, original) => EncodeLikeOriginal(replacements[path], original),
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

        EpubResource[] htmlResources = document.Resources.Where(IsHtmlResource).ToArray();
        if (htmlResources.Length == 0) throw new InvalidDataException("The EPUB manifest contains no local HTML or XHTML content documents.");
        var resourcesByUri = new Dictionary<string, EpubResource>(StringComparer.Ordinal);
        foreach (EpubResource resource in document.Resources.Where(item => !item.IsRemote && item.Data != null)) {
            resourcesByUri[CreatePackageUri(resource.Path).AbsoluteUri] = resource;
        }

        var parts = new List<HtmlContentSafetyPackagePart>(htmlResources.Length);
        foreach (EpubResource resource in htmlResources) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = resource.Data ?? throw new InvalidDataException(
                "EPUB content document payload was not retained within the configured limits: " + resource.Path);
            if (resource.Encryption?.RequiresDecryption == true) {
                throw new InvalidDataException("EPUB content document requires unsupported decryption: " + resource.Path);
            }
            string html = OfficeContentSafetyInputGuard.DecodeText(bytes, options);
            var renderOptions = CreateEpubRenderOptions(resource.Path, resourcesByUri, options);
            bool xhtml = IsXhtmlResource(resource);
            parts.Add(new HtmlContentSafetyPackagePart(
                resource.Path,
                html,
                "EPUB/" + resource.Path,
                renderOptions,
                serializeAsXhtml: xhtml));
        }
        return new EpubContentSafetyPackage(document, parts);
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
        IReadOnlyDictionary<string, EpubResource> resourcesByUri,
        OfficeContentSafetyOptions options) {
        var resourcePolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile();
        resourcePolicy.AllowedUrlSchemes.Add("epub");
        var renderOptions = new HtmlRenderOptions {
            BaseUri = CreatePackageUri(contentPath),
            ResourceUrlPolicy = resourcePolicy,
            MaxInputCharacters = options.MaxCharacters,
            MaxResourceBytes = options.MaxInputBytes,
            MaxTotalResourceBytes = options.MaxExpandedPackageBytes,
            MaxResourceCount = options.MaxPackageEntries,
            MaxResourceRequests = options.MaxPackageEntries
        };
        renderOptions.ResourceResolver = (request, token) => {
            token.ThrowIfCancellationRequested();
            if (!request.Uri.Scheme.Equals("epub", StringComparison.OrdinalIgnoreCase)
                || !resourcesByUri.TryGetValue(request.Uri.AbsoluteUri, out EpubResource? resource)
                || resource.Data == null) {
                return Task.FromResult<HtmlResolvedResource?>(null);
            }
            return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                resource.Data,
                string.IsNullOrWhiteSpace(resource.MediaType) ? "application/octet-stream" : resource.MediaType!));
        };
        return renderOptions;
    }

    private static Uri CreatePackageUri(string path) {
        string escaped = string.Join("/", path.Split('/').Select(Uri.EscapeDataString));
        return new Uri("epub://package/" + escaped, UriKind.Absolute);
    }

    private static bool IsHtmlResource(EpubResource resource) => !resource.IsRemote && (
        string.Equals(resource.MediaType, "application/xhtml+xml", StringComparison.OrdinalIgnoreCase)
        || string.Equals(resource.MediaType, "text/html", StringComparison.OrdinalIgnoreCase)
        || resource.Path.EndsWith(".xhtml", StringComparison.OrdinalIgnoreCase)
        || resource.Path.EndsWith(".html", StringComparison.OrdinalIgnoreCase)
        || resource.Path.EndsWith(".htm", StringComparison.OrdinalIgnoreCase));

    private static bool IsXhtmlResource(EpubResource resource) =>
        string.Equals(resource.MediaType, "application/xhtml+xml", StringComparison.OrdinalIgnoreCase)
        || resource.Path.EndsWith(".xhtml", StringComparison.OrdinalIgnoreCase);

    private static void ThrowForIncompleteContentSafetyRead(EpubDocument document) {
        string[] blockingPrefixes = {
            "epub.archive.unsafe-path",
            "epub.archive.duplicate-path",
            "epub.container.rootfile-fallback",
            "epub.container.rootfile-path-invalid",
            "epub.container.rootfile-duplicate",
            "epub.container.rootfile-missing",
            "epub.container.multiple-rootfiles",
            "epub.package.missing",
            "epub.package.metadata-size-limit",
            "epub.package.invalid-xml",
            "epub.manifest.invalid-path",
            "epub.manifest.duplicate-id",
            "epub.manifest.duplicate-target",
            "epub.resource.count-limit",
            "epub.resource.missing",
            "epub.resource.size-limit",
            "epub.resource.total-size-limit",
            "epub.resource.encrypted",
            "epub.encryption.invalid-xml",
            "epub.encryption.metadata-size-limit",
            "epub.encryption.resource-path-invalid",
            "epub.encryption.duplicate-resource",
            "epub.encryption.resource-missing",
            "epub.encryption.unsupported"
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
        var caseFolded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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
            if (!caseFolded.Add(identity)) {
                throw new InvalidDataException("EPUB package contains case-colliding entry paths that cannot be resolved unambiguously: " + identity);
            }
        }
    }

    private static byte[] EncodeLikeOriginal(string text, byte[] original) {
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
        } else {
            encoding = new UTF8Encoding(false, true);
            preamble = Array.Empty<byte>();
        }
        byte[] payload = encoding.GetBytes(text);
        if (preamble.Length == 0) return payload;
        var result = new byte[preamble.Length + payload.Length];
        Buffer.BlockCopy(preamble, 0, result, 0, preamble.Length);
        Buffer.BlockCopy(payload, 0, result, preamble.Length, payload.Length);
        return result;
    }

    private sealed class EpubContentSafetyPackage {
        internal EpubContentSafetyPackage(
            EpubDocument document,
            IReadOnlyList<HtmlContentSafetyPackagePart> parts) {
            Document = document;
            Parts = parts;
        }

        internal EpubDocument Document { get; }
        internal IReadOnlyList<HtmlContentSafetyPackagePart> Parts { get; }
    }
}
