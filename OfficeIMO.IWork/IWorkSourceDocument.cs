using OfficeIMO.IWork.Internal;
using System.Text;
using System.Xml;

namespace OfficeIMO.IWork;

/// <summary>Bounded, read-only representation of a modern Pages, Numbers, or Keynote source package.</summary>
public sealed partial class IWorkSourceDocument {
    private readonly IWorkReadOptions _options;
    private readonly IWorkObjectIndex _index;

    private IWorkSourceDocument(IWorkDocumentKind kind, IWorkPackageData package,
        IReadOnlyList<IWorkArchiveRecord> records, IWorkReadOptions options) {
        Kind = kind;
        ContainerKind = package.ContainerKind;
        Entries = Array.AsReadOnly(package.Entries.ToArray());
        Records = Array.AsReadOnly(records.ToArray());
        _options = options;
        _index = new IWorkObjectIndex(Records, options);
        BuildVersions = Array.AsReadOnly(ReadBuildVersions(Entries).ToArray());
        Previews = Array.AsReadOnly(ReadPreviews(Entries, options.MaximumPackageBytes).ToArray());
        Diagnostics = Array.AsReadOnly(new[] {
            new IWorkDiagnostic(IWorkDiagnosticSeverity.Information, "IWORK_SOURCE_READ",
                $"Read {records.Count} IWA payload records from {package.Entries.Count} package entries.")
        });
    }

    /// <summary>Gets the source application.</summary>
    public IWorkDocumentKind Kind { get; }
    /// <summary>Gets the physical source layout.</summary>
    public IWorkContainerKind ContainerKind { get; }
    /// <summary>Gets preserved package entries, including resources and metadata.</summary>
    public IReadOnlyList<IWorkPackageEntry> Entries { get; }
    /// <summary>Gets every preserved IWA payload record, including auxiliary payloads.</summary>
    public IReadOnlyList<IWorkArchiveRecord> Records { get; }
    /// <summary>Gets build-history strings recorded by the producer.</summary>
    public IReadOnlyList<string> BuildVersions { get; }
    /// <summary>Gets embedded raster or PDF previews ordered from most useful to least useful.</summary>
    public IReadOnlyList<IWorkPreviewAsset> Previews { get; }
    /// <summary>Gets the snapshotted projection mode requested when the source was opened.</summary>
    public IWorkImportMode RequestedImportMode => _options.ImportMode;
    /// <summary>Gets package-level diagnostics.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
    /// <summary>Gets the preferred raster asset for a visual fallback.</summary>
    public IWorkPreviewAsset? PreferredRasterPreview => Previews.FirstOrDefault(preview =>
        preview.MediaType == "image/jpeg" || preview.MediaType == "image/png");

    /// <summary>Opens an iWork file or directory bundle and detects its application kind.</summary>
    public static IWorkSourceDocument Open(string path, IWorkReadOptions? options = null) =>
        OpenPath(path, expectedKind: null, options);

    /// <summary>Opens an iWork file or directory bundle and verifies the expected application kind.</summary>
    public static IWorkSourceDocument Open(string path, IWorkDocumentKind expectedKind,
        IWorkReadOptions? options = null) => OpenPath(path, expectedKind, options);

    /// <summary>Opens a ZIP-based iWork stream using the supplied application kind.</summary>
    public static IWorkSourceDocument Open(Stream stream, IWorkDocumentKind kind,
        IWorkReadOptions? options = null) {
        IWorkReadOptions resolved = (options ?? new IWorkReadOptions()).Snapshot();
        IWorkPackageData package = IWorkContainerReader.Read(stream, resolved);
        return Create(package, kind, resolved, expectedKind: kind);
    }

    internal IWorkObjectIndex Index => _index;
    internal IWorkReadOptions Options => _options;

    internal IWorkImportReport CreateReport(IWorkProjectionKind projectionKind,
        IReadOnlyList<IWorkDiagnostic> projectionDiagnostics,
        IWorkPreviewAsset? preview, int reconstructedItemCount) {
        IWorkArchiveRecord[] allUnsupported = Records.ToArray();
        IReadOnlyList<IWorkArchiveRecord> unsupported = _options.PreserveUnsupportedRecords
            ? allUnsupported
            : Array.Empty<IWorkArchiveRecord>();
        return new IWorkImportReport(
            Kind,
            projectionKind,
            BuildVersions,
            unsupported,
            Diagnostics.Concat(projectionDiagnostics).ToArray(),
            preview,
            Records.Count,
            allUnsupported.Length,
            reconstructedItemCount);
    }

    private static IWorkSourceDocument OpenPath(string path, IWorkDocumentKind? expectedKind,
        IWorkReadOptions? options) {
        IWorkReadOptions resolved = (options ?? new IWorkReadOptions()).Snapshot();
        IWorkPackageData package = IWorkContainerReader.Read(path, resolved);
        IWorkDocumentKind? extensionKind = KindFromExtension(Path.GetExtension(path));
        return Create(package, expectedKind ?? extensionKind, resolved, expectedKind);
    }

    private static IWorkSourceDocument Create(IWorkPackageData package, IWorkDocumentKind? hint,
        IWorkReadOptions options, IWorkDocumentKind? expectedKind) {
        if (!package.Entries.Any(entry => IWorkArchiveParser.IsIndexArchivePath(entry.Path))) {
            string[] legacyMarkers = { "index.xml", "index.apxl", "index.apxl.gz" };
            if (package.Entries.Any(entry => legacyMarkers.Contains(entry.Path, StringComparer.OrdinalIgnoreCase))) {
                throw new NotSupportedException("Pre-2013 iWork packages are not supported; this reader requires IWA archives.");
            }
            throw new InvalidDataException("The package does not contain modern iWork IWA archives.");
        }

        IReadOnlyList<IWorkArchiveRecord> records = IWorkArchiveParser.Parse(package.Entries, options);
        IWorkDocumentKind detected = DetectKind(package.Entries, records, hint);
        if (expectedKind.HasValue && expectedKind.Value != detected) {
            throw new InvalidDataException($"The package is {detected}, not the expected {expectedKind.Value} source.");
        }
        return new IWorkSourceDocument(detected, package, records, options);
    }

    private static IWorkDocumentKind DetectKind(IReadOnlyList<IWorkPackageEntry> entries,
        IReadOnlyList<IWorkArchiveRecord> records, IWorkDocumentKind? hint) {
        bool hasSlides = entries.Any(entry => IsKeynoteSlideArchive(entry.Path));
        if (hasSlides) return IWorkDocumentKind.Keynote;
        if (records.Any(record => record.IsPrimary && record.MessageType == 10000)) return IWorkDocumentKind.Pages;
        if (records.Any(record => record.IsPrimary && record.MessageType == 2)
            && records.Any(record => record.IsPrimary && record.MessageType is 6000 or 6001 or 6002)) {
            return IWorkDocumentKind.Numbers;
        }
        if (hint.HasValue) return hint.Value;
        if (records.Any(record => record.IsPrimary && record.MessageType == 1)) return IWorkDocumentKind.Numbers;
        throw new InvalidDataException("The iWork application kind could not be identified from the package structure.");
    }

    private static bool IsKeynoteSlideArchive(string path) {
        const string indexPrefix = "Index/";
        if (!path.StartsWith(indexPrefix, StringComparison.OrdinalIgnoreCase)
            || path.IndexOf('/', indexPrefix.Length) >= 0) return false;
        string name = path.Substring(indexPrefix.Length);
        return IsArchiveName(name, "Slide")
            || IsArchiveName(name, "MasterSlide")
            || IsArchiveName(name, "TemplateSlide");
    }

    private static bool IsArchiveName(string name, string stem) =>
        name.Equals(stem + ".iwa", StringComparison.OrdinalIgnoreCase)
        || name.StartsWith(stem + "-", StringComparison.OrdinalIgnoreCase)
        && name.EndsWith(".iwa", StringComparison.OrdinalIgnoreCase);

    private static IWorkDocumentKind? KindFromExtension(string extension) => extension.ToLowerInvariant() switch {
        ".pages" => IWorkDocumentKind.Pages,
        ".numbers" => IWorkDocumentKind.Numbers,
        ".key" => IWorkDocumentKind.Keynote,
        _ => null
    };

    private static IReadOnlyList<string> ReadBuildVersions(IReadOnlyList<IWorkPackageEntry> entries) {
        IWorkPackageEntry? entry = entries.FirstOrDefault(candidate =>
            string.Equals(candidate.Path, "Metadata/BuildVersionHistory.plist", StringComparison.OrdinalIgnoreCase));
        if (entry == null) return Array.Empty<string>();
        try {
            using var stream = new MemoryStream(entry.Bytes, writable: false);
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Ignore,
                XmlResolver = null,
                MaxCharactersInDocument = entry.Length,
                MaxCharactersFromEntities = 0
            };
            using XmlReader reader = XmlReader.Create(stream, settings);
            int nodeCount = 0;
            int maximumNodes = checked((int)Math.Min(100_000L, Math.Max(16L, entry.Length)));
            long maximumCharacters = Math.Min(16L * 1024 * 1024, entry.Length);
            long extractedCharacters = 0;
            bool sawPlist = false;
            bool sawArray = false;
            var values = new List<string>();
            while (ReadBuildHistoryNode(reader, ref nodeCount, maximumNodes)) {
                if (reader.Depth > 8) return Array.Empty<string>();
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.Depth == 0 && reader.LocalName == "plist" && reader.NamespaceURI.Length == 0
                    && !sawPlist) {
                    sawPlist = true;
                    continue;
                }
                if (reader.Depth == 1 && reader.LocalName == "array" && reader.NamespaceURI.Length == 0
                    && sawPlist && !sawArray) {
                    sawArray = true;
                    continue;
                }
                if (reader.Depth != 2 || reader.LocalName != "string" || reader.NamespaceURI.Length != 0
                    || !sawArray
                    || !TryReadBuildHistoryString(reader, ref nodeCount, maximumNodes,
                        ref extractedCharacters, maximumCharacters, out string value)) {
                    return Array.Empty<string>();
                }
                if (!string.IsNullOrWhiteSpace(value)) values.Add(value);
            }
            return sawPlist && sawArray ? values : Array.Empty<string>();
        } catch (Exception exception) when (exception is XmlException or InvalidOperationException
                or InvalidDataException or IOException) {
            return Array.Empty<string>();
        }
    }

    private static bool ReadBuildHistoryNode(XmlReader reader, ref int nodeCount, int maximumNodes) {
        if (!reader.Read()) return false;
        if (nodeCount >= maximumNodes) throw new InvalidDataException("Build history XML exceeds the node limit.");
        nodeCount++;
        return true;
    }

    private static bool TryReadBuildHistoryString(XmlReader reader, ref int nodeCount,
        int maximumNodes, ref long extractedCharacters, long maximumCharacters,
        out string value) {
        value = string.Empty;
        if (reader.IsEmptyElement) return true;
        var builder = new StringBuilder();
        var buffer = new char[1024];
        while (ReadBuildHistoryNode(reader, ref nodeCount, maximumNodes)) {
            if (reader.Depth > 8) return false;
            if (reader.NodeType == XmlNodeType.EndElement) {
                if (reader.Depth != 2 || reader.LocalName != "string" || reader.NamespaceURI.Length != 0) {
                    return false;
                }
                value = builder.ToString();
                return true;
            }
            if (reader.NodeType is not (XmlNodeType.Text or XmlNodeType.CDATA
                    or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace)) {
                return false;
            }
            int read;
            while ((read = reader.ReadValueChunk(buffer, 0, buffer.Length)) > 0) {
                if (extractedCharacters > maximumCharacters - read) return false;
                extractedCharacters += read;
                builder.Append(buffer, 0, read);
            }
        }
        return false;
    }

    private static IReadOnlyList<IWorkPreviewAsset> ReadPreviews(
        IReadOnlyList<IWorkPackageEntry> entries, long maximumDecodedBytes) {
        var previews = new List<IWorkPreviewAsset>();
        var recognizedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long remainingDecodedBytes = maximumDecodedBytes;
        foreach (IWorkPackageEntry entry in entries) {
            string lower = entry.Path.ToLowerInvariant();
            string? mediaType = lower.EndsWith(".jpg", StringComparison.Ordinal) || lower.EndsWith(".jpeg", StringComparison.Ordinal)
                ? "image/jpeg"
                : lower.EndsWith(".png", StringComparison.Ordinal)
                    ? "image/png"
                    : lower.EndsWith(".pdf", StringComparison.Ordinal)
                        ? "application/pdf"
                        : null;
            if (mediaType == null || !IsKnownPreviewPath(lower)
                || !recognizedPaths.Add(entry.Path)
                || !HasExpectedSignature(entry.Bytes, mediaType)) continue;
            IWorkVisualCoverage coverage = mediaType == "application/pdf"
                ? IWorkVisualCoverage.FullDocument
                : IWorkVisualCoverage.FirstPageOrCompositePreview;
            (int? width, int? height) = IWorkImageInfo.Read(
                entry.Bytes, mediaType, remainingDecodedBytes, out long decodedBytes);
            if (mediaType != "application/pdf" && (!width.HasValue || !height.HasValue)) continue;
            if (decodedBytes < 0 || decodedBytes > remainingDecodedBytes) continue;
            remainingDecodedBytes -= decodedBytes;
            previews.Add(new IWorkPreviewAsset(entry.Path, mediaType, coverage, width, height, entry.Bytes));
        }
        return previews
            .OrderBy(preview => preview.MediaType == "application/pdf" ? 0 : 1)
            .ThenBy(preview => PreviewRank(preview.Path))
            .ThenByDescending(preview => preview.Length)
            .ToArray();
    }

    private static int PreviewRank(string path) {
        string lower = path.ToLowerInvariant();
        if (lower.EndsWith("preview.jpg", StringComparison.Ordinal) || lower.EndsWith("preview.png", StringComparison.Ordinal)) return 0;
        if (lower.Contains("preview-web")) return 1;
        if (lower.Contains("preview-micro")) return 3;
        return 2;
    }

    private static bool IsKnownPreviewPath(string lowerPath) {
        if (lowerPath.IndexOf('/') < 0) {
            return lowerPath is "preview.jpg" or "preview.jpeg" or "preview.png" or "preview.pdf"
                or "preview-web.jpg" or "preview-web.jpeg" or "preview-web.png"
                or "preview-micro.jpg" or "preview-micro.jpeg" or "preview-micro.png";
        }
        return lowerPath is "quicklook/preview.pdf" or "quicklook/preview.jpg" or "quicklook/preview.jpeg"
            or "quicklook/preview.png" or "quicklook/thumbnail.jpg" or "quicklook/thumbnail.jpeg"
            or "quicklook/thumbnail.png";
    }

    private static bool HasExpectedSignature(byte[] bytes, string mediaType) {
        if (mediaType == "image/jpeg") return bytes.Length >= 3 && bytes[0] == 0xff && bytes[1] == 0xd8 && bytes[2] == 0xff;
        if (mediaType == "image/png") {
            byte[] signature = { 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a };
            return bytes.Length >= signature.Length && signature.Where((value, index) => bytes[index] != value).Any() == false;
        }
        return mediaType == "application/pdf" && IWorkPdfInfo.IsComplete(bytes);
    }
}
