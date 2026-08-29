using OfficeIMO.IWork.Internal;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.IWork;

/// <summary>Bounded, read-only representation of a modern Pages, Numbers, or Keynote source package.</summary>
public sealed partial class IWorkSourceDocument {
    private readonly IWorkReadOptions _options;
    private readonly IWorkObjectIndex _index;

    private IWorkSourceDocument(IWorkDocumentKind kind, IWorkPackageData package,
        IReadOnlyList<IWorkArchiveRecord> records, IWorkReadOptions options) {
        Kind = kind;
        ContainerKind = package.ContainerKind;
        Entries = package.Entries;
        Records = records;
        _options = options;
        _index = new IWorkObjectIndex(records, options);
        BuildVersions = ReadBuildVersions(package.Entries);
        Previews = ReadPreviews(package.Entries);
        Diagnostics = new[] {
            new IWorkDiagnostic(IWorkDiagnosticSeverity.Information, "IWORK_SOURCE_READ",
                $"Read {records.Count} IWA payload records from {package.Entries.Count} package entries.")
        };
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
        IReadOnlyCollection<ulong> recognizedIdentifiers, IReadOnlyList<IWorkDiagnostic> projectionDiagnostics,
        IWorkPreviewAsset? preview, int reconstructedItemCount) {
        HashSet<ulong> recognized = new(recognizedIdentifiers);
        IWorkArchiveRecord[] allUnsupported = Records
            .Where(record => !record.IsPrimary || !recognized.Contains(record.Identifier))
            .ToArray();
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
        return Create(package, extensionKind ?? expectedKind, resolved, expectedKind);
    }

    private static IWorkSourceDocument Create(IWorkPackageData package, IWorkDocumentKind? hint,
        IWorkReadOptions options, IWorkDocumentKind? expectedKind) {
        if (!package.Entries.Any(entry => entry.Path.EndsWith(".iwa", StringComparison.OrdinalIgnoreCase))) {
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
        bool hasSlides = entries.Any(entry =>
            entry.Path.StartsWith("Index/Slide", StringComparison.OrdinalIgnoreCase)
            || entry.Path.StartsWith("Index/MasterSlide", StringComparison.OrdinalIgnoreCase)
            || entry.Path.StartsWith("Index/TemplateSlide", StringComparison.OrdinalIgnoreCase));
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
            XDocument document = XDocument.Load(reader, LoadOptions.None);
            return document.Descendants("string")
                .Select(element => element.Value)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToArray();
        } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidOperationException) {
            return Array.Empty<string>();
        }
    }

    private static IReadOnlyList<IWorkPreviewAsset> ReadPreviews(IReadOnlyList<IWorkPackageEntry> entries) {
        var previews = new List<IWorkPreviewAsset>();
        foreach (IWorkPackageEntry entry in entries) {
            string lower = entry.Path.ToLowerInvariant();
            string? mediaType = lower.EndsWith(".jpg", StringComparison.Ordinal) || lower.EndsWith(".jpeg", StringComparison.Ordinal)
                ? "image/jpeg"
                : lower.EndsWith(".png", StringComparison.Ordinal)
                    ? "image/png"
                    : lower.EndsWith(".pdf", StringComparison.Ordinal)
                        ? "application/pdf"
                        : null;
            if (mediaType == null || !IsKnownPreviewPath(lower) || !HasExpectedSignature(entry.Bytes, mediaType)) continue;
            IWorkVisualCoverage coverage = mediaType == "application/pdf"
                ? IWorkVisualCoverage.FullDocument
                : IWorkVisualCoverage.FirstPageOrCompositePreview;
            (int? width, int? height) = IWorkImageInfo.Read(entry.Bytes, mediaType);
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
        return mediaType == "application/pdf" && bytes.Length >= 5
            && bytes[0] == (byte)'%' && bytes[1] == (byte)'P' && bytes[2] == (byte)'D'
            && bytes[3] == (byte)'F' && bytes[4] == (byte)'-';
    }
}