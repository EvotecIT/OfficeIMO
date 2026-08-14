using OfficeIMO.Provenance;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Epub;

public sealed partial class EpubDocument {
    /// <summary>Inspects C2PA and IPTC provenance in an EPUB package and its supported embedded images.</summary>
    public static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        options ??= new OfficeProvenanceOptions();
        string fullPath = Path.GetFullPath(filePath);
        byte[] data;
        using (Stream stream = File.OpenRead(fullPath)) data = OfficeProvenanceBinary.ReadBounded(stream, options.MaxAssetBytes);
        ValidatePackage(data, options);
        return OfficeProvenanceInspector.Inspect(data, fullPath, options);
    }

    /// <summary>Removes selected provenance and atomically writes an EPUB package.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(
            inputPath,
            outputPath,
            options,
            StripPackageSignatures,
            HasPackageSignatures,
            ValidatePackage,
            removeOpcManifestReferences: false,
            validateOpcMetadata: false);

    /// <summary>Removes selected provenance from encoded EPUB package bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] packageBytes,
        string fileName = "publication.epub",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(
            packageBytes,
            fileName,
            options,
            StripPackageSignatures,
            HasPackageSignatures,
            ValidatePackage,
            removeOpcManifestReferences: false,
            validateOpcMetadata: false);

    private static void ValidatePackage(byte[] data, OfficeProvenanceOptions options) {
        OfficeProvenanceZip.ValidateMimetypeEntry(data, "application/epub+zip", options.MaxContainerEntries);
        using var input = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
        ZipArchiveEntry[] containers = archive.Entries
            .Where(entry => entry.FullName.Replace('\\', '/').Equals("META-INF/container.xml", StringComparison.Ordinal))
            .ToArray();
        if (containers.Length != 1) throw new InvalidDataException("The EPUB package must contain exactly one META-INF/container.xml part.");
        using Stream containerStream = containers[0].Open();
        byte[] containerXml = OfficeProvenanceBinary.ReadBounded(containerStream, options.MaxAssetBytes);
        if (containerXml.LongLength > options.MaxExpandedContainerBytes) {
            throw new InvalidDataException("The EPUB container metadata exceeds the configured expanded-container limit.");
        }
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(containerXml, options, "EPUB container metadata");
        using var xmlInput = new MemoryStream(containerXml, writable: false);
        using XmlReader reader = XmlReader.Create(xmlInput, OfficeProvenanceXml.CreateReaderSettings(options));
        XDocument document = XDocument.Load(reader, LoadOptions.None);
        XNamespace containerNamespace = "urn:oasis:names:tc:opendocument:xmlns:container";
        if (document.Root?.Name != containerNamespace + "container") {
            throw new InvalidDataException("The EPUB container metadata has an unexpected root element.");
        }
        Dictionary<string, ZipArchiveEntry[]> entriesByPath = archive.Entries
            .GroupBy(entry => entry.FullName.Replace('\\', '/'), StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        XElement[] rootfileContainers = document.Root.Elements(containerNamespace + "rootfiles").ToArray();
        if (rootfileContainers.Length != 1) {
            throw new InvalidDataException("The EPUB container metadata must contain exactly one rootfiles element.");
        }
        int declaredRootfiles = 0;
        bool hasValidPackageDocument = false;
        long expandedMetadataBytes = containerXml.LongLength;
        foreach (XElement rootfile in rootfileContainers[0].Elements(containerNamespace + "rootfile")) {
            if (++declaredRootfiles > options.MaxContainerEntries) {
                throw new InvalidDataException("The EPUB container declares too many rootfiles.");
            }
            if (!TryNormalizeRootfilePath((string?)rootfile.Attribute("full-path"), out string path)) {
                throw new InvalidDataException("The EPUB container declares an invalid rootfile path.");
            }
            if (entriesByPath.TryGetValue(path, out ZipArchiveEntry[]? matches)) {
                if (matches.Length != 1) throw new InvalidDataException("The EPUB package contains duplicate declared rootfile entries.");
                if (string.Equals(
                    (string?)rootfile.Attribute("media-type"),
                    "application/oebps-package+xml",
                    StringComparison.Ordinal) &&
                    TryValidateOpfPackage(matches[0], options, ref expandedMetadataBytes)) {
                    hasValidPackageDocument = true;
                }
            }
        }
        if (declaredRootfiles == 0 || !hasValidPackageDocument) {
            throw new InvalidDataException("The EPUB container must declare at least one bounded OPF package document that exists in the package.");
        }
    }

    private static bool TryValidateOpfPackage(
        ZipArchiveEntry entry,
        OfficeProvenanceOptions options,
        ref long expandedMetadataBytes) {
        long remainingExpandedBytes = options.MaxExpandedContainerBytes - expandedMetadataBytes;
        if (remainingExpandedBytes < 0 || entry.Length > remainingExpandedBytes) {
            throw new InvalidDataException("The EPUB package metadata exceeds the configured expanded-container limit.");
        }
        using Stream source = entry.Open();
        byte[] opf = OfficeProvenanceBinary.ReadBounded(source, Math.Min(options.MaxAssetBytes, remainingExpandedBytes));
        expandedMetadataBytes += opf.LongLength;
        try {
            OfficeProvenanceXml.ValidateMaterializedNodeBudget(opf, options, "EPUB package document");
            using var input = new MemoryStream(opf, writable: false);
            using XmlReader reader = XmlReader.Create(input, OfficeProvenanceXml.CreateReaderSettings(options));
            XDocument document = XDocument.Load(reader, LoadOptions.None);
            XNamespace opfNamespace = "http://www.idpf.org/2007/opf";
            return document.Root?.Name == opfNamespace + "package";
        } catch (XmlException) {
            return false;
        }
    }

    private static bool TryNormalizeRootfilePath(string? value, out string normalized) {
        normalized = string.Empty;
        if (string.IsNullOrWhiteSpace(value) || value!.IndexOfAny(new[] { '\\', '?', '#' }) >= 0 || value.StartsWith("/", StringComparison.Ordinal)) return false;
        var segments = new List<string>();
        foreach (string encodedSegment in value.Split('/')) {
            if (encodedSegment.Length == 0) return false;
            string segment;
            try { segment = Uri.UnescapeDataString(encodedSegment); }
            catch (UriFormatException) { return false; }
            if (segment.IndexOfAny(new[] { '/', '\\', '?', '#' }) >= 0 || segment.Length == 0) return false;
            if (segment == ".") continue;
            if (segment == "..") {
                if (segments.Count == 0) return false;
                segments.RemoveAt(segments.Count - 1);
                continue;
            }
            segments.Add(segment);
        }
        if (segments.Count == 0) return false;
        normalized = string.Join("/", segments);
        return true;
    }

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions _) => OfficeProvenanceZip.HasEntry(data, path =>
        path.Equals("META-INF/signatures.xml", StringComparison.Ordinal));

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceOptions limits) {
        return OfficeProvenanceZip.RemoveEntries(
            data,
            path => path.Equals("META-INF/signatures.xml", StringComparison.Ordinal),
            limits.MaxExpandedContainerBytes);
    }
}
