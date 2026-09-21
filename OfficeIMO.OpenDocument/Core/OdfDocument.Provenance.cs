using OfficeIMO.Provenance;
using OfficeIMO.Core.Internal;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    private const string ProvenanceManifestPath = "META-INF/content_credential.c2pa";
    /// <summary>Inspects C2PA and IPTC provenance in an ODF package and its supported embedded images.</summary>
    public static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        options ??= new OfficeProvenanceOptions();
        string fullPath = Path.GetFullPath(filePath);
        byte[] data;
        using (Stream stream = File.OpenRead(fullPath)) data = OfficeProvenanceBinary.ReadBounded(stream, options.MaxAssetBytes, options.CancellationToken);
        ValidatePackageForInspection(data, fullPath, options);
        return OfficeProvenanceInspector.Inspect(data, fullPath, options);
    }

    /// <summary>Removes selected provenance and atomically writes an ODF package.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeProvenanceRemovalOptions();
        byte[] data;
        using (Stream stream = File.OpenRead(Path.GetFullPath(inputPath))) data = OfficeProvenanceBinary.ReadBounded(stream, options.Limits.MaxAssetBytes, options.Limits.CancellationToken);
        OfficeProvenanceRemovalResult result = RemoveProvenance(data, Path.GetFileName(inputPath), options);
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), result.ToArray());
        return result;
    }

    /// <summary>Removes selected provenance from encoded ODF package bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] packageBytes,
        string fileName = "document.odt",
        OfficeProvenanceRemovalOptions? options = null) {
        options ??= new OfficeProvenanceRemovalOptions();
        return OfficeProvenancePackageMutation.Remove(
            packageBytes,
            fileName,
            options,
            StripPackageSignatures,
            HasPackageSignatures,
            ValidatePackage,
            removeOpcManifestReferences: false,
            validateOpcMetadata: false,
            shouldReplacePackageMetadata: path => path == "META-INF/manifest.xml",
            replacePackageMetadata: (_, manifest, nativeManifestRemoved) =>
                nativeManifestRemoved
                    ? RemoveManifestEntries(
                        manifest,
                        options.Limits,
                        path => path == ProvenanceManifestPath)
                    : manifest);
    }

    private static readonly string[] SupportedMimetypes = {
        OdfMediaTypes.Text,
        OdfMediaTypes.Spreadsheet,
        OdfMediaTypes.Presentation,
        OdfMediaTypes.Graphics,
        OdfMediaTypes.TextTemplate,
        OdfMediaTypes.SpreadsheetTemplate,
        OdfMediaTypes.PresentationTemplate,
        OdfMediaTypes.GraphicsTemplate
    };

    private static void ValidatePackage(byte[] data, string fileName, OfficeProvenanceOptions options) {
        ValidatePackage(data, fileName, options, rejectEncrypted: true);
    }

    private static void ValidatePackageForInspection(byte[] data, string fileName, OfficeProvenanceOptions options) {
        ValidatePackage(data, fileName, options, rejectEncrypted: false);
    }

    private static void ValidatePackage(byte[] data, string fileName, OfficeProvenanceOptions options, bool rejectEncrypted) {
        string mediaType = OfficeProvenanceZip.ReadValidatedMimetypeEntry(data, SupportedMimetypes, options.MaxContainerEntries);
        string expectedMediaType = Path.GetExtension(fileName).ToLowerInvariant() switch {
            ".odt" => OdfMediaTypes.Text,
            ".ods" => OdfMediaTypes.Spreadsheet,
            ".odp" => OdfMediaTypes.Presentation,
            ".odg" => OdfMediaTypes.Graphics,
            ".ott" => OdfMediaTypes.TextTemplate,
            ".ots" => OdfMediaTypes.SpreadsheetTemplate,
            ".otp" => OdfMediaTypes.PresentationTemplate,
            ".otg" => OdfMediaTypes.GraphicsTemplate,
            _ => throw new NotSupportedException("The filename extension is not an OfficeIMO-owned OpenDocument format.")
        };
        if (!string.Equals(mediaType, expectedMediaType, StringComparison.Ordinal)) {
            throw new InvalidDataException($"The OpenDocument package media type '{mediaType}' does not match filename extension '{Path.GetExtension(fileName)}'.");
        }
        using var input = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
        var exactEntryNames = new HashSet<string>(StringComparer.Ordinal);
        var foldedEntryNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (ZipArchiveEntry entry in archive.Entries) {
            options.CancellationToken.ThrowIfCancellationRequested();
            string normalized = OfficeArchiveSafety.NormalizeEntryName(entry.FullName);
            if (!string.Equals(normalized, entry.FullName, StringComparison.Ordinal) ||
                OfficeArchiveSafety.IsUnsafePath(normalized)) {
                throw new InvalidDataException($"OpenDocument package contains unsafe or non-canonical entry path '{entry.FullName}'.");
            }
            if (!exactEntryNames.Add(normalized) || !foldedEntryNames.Add(normalized)) {
                throw new InvalidDataException($"OpenDocument package contains duplicate or case-ambiguous entry '{normalized}'.");
            }
        }
        ZipArchiveEntry[] manifestEntries = archive.Entries
            .Where(entry => string.Equals(entry.FullName, "META-INF/manifest.xml", StringComparison.Ordinal))
            .ToArray();
        if (manifestEntries.Length != 1) throw new InvalidDataException("OpenDocument package must contain exactly one 'META-INF/manifest.xml'.");
        int contentEntryCount = archive.Entries.Count(entry =>
            string.Equals(entry.FullName, "content.xml", StringComparison.Ordinal));
        if (contentEntryCount != 1) throw new InvalidDataException("OpenDocument package must contain exactly one 'content.xml'.");

        long maximumManifestBytes = Math.Min(options.MaxAssetBytes, options.MaxExpandedContainerBytes);
        byte[] manifestBytes;
        using (Stream stream = manifestEntries[0].Open()) manifestBytes = OfficeProvenanceBinary.ReadBounded(stream, maximumManifestBytes, options.CancellationToken);
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(manifestBytes, options, "ODF manifest");

        XDocument manifest;
        using (var stream = new MemoryStream(manifestBytes, writable: false))
        using (XmlReader reader = XmlReader.Create(stream, OfficeProvenanceXml.CreateReaderSettings(options))) {
            manifest = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        options.CancellationToken.ThrowIfCancellationRequested();
        XNamespace manifestNamespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        XElement root = manifest.Root ?? throw new InvalidDataException("OpenDocument manifest has no root element.");
        if (root.Name != manifestNamespace + "manifest") {
            throw new InvalidDataException("OpenDocument manifest root must be 'manifest:manifest'.");
        }
        if (rejectEncrypted && root.Descendants(manifestNamespace + "encryption-data").Any()) {
            throw new OdfEncryptedPackageException("Encrypted OpenDocument packages are detected but not supported for provenance mutation.");
        }

        XElement[] packageRoots = root.Elements(manifestNamespace + "file-entry")
            .Where(element => string.Equals((string?)element.Attribute(manifestNamespace + "full-path"), "/", StringComparison.Ordinal))
            .ToArray();
        if (packageRoots.Length != 1) throw new InvalidDataException("OpenDocument manifest must contain exactly one package root entry.");
        string? manifestMediaType = (string?)packageRoots[0].Attribute(manifestNamespace + "media-type");
        if (!string.Equals(mediaType, manifestMediaType, StringComparison.Ordinal)) {
            throw new InvalidDataException("OpenDocument mimetype does not match the root manifest media type.");
        }
    }

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) =>
        FindSignatureEntries(data, options.Limits).Count > 0;

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) {
        OfficeProvenanceOptions limits = options.Limits;
        HashSet<string> signatureEntries = FindSignatureEntries(data, limits);
        return OfficeProvenanceZip.RemoveEntries(
            data,
            signatureEntries.Contains,
            limits.MaxExpandedContainerBytes,
            path => path == "META-INF/manifest.xml",
            (_, manifest) => RemoveManifestEntries(
                manifest,
                limits,
                signatureEntries.Contains),
            limits.MaxAssetBytes,
            options.EffectiveMaxOutputBytes,
            limits.CancellationToken);
    }

    internal static HashSet<string> FindSignatureEntries(byte[] data, OfficeProvenanceOptions limits) {
        var signatures = new HashSet<string>(StringComparer.Ordinal);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        Dictionary<ZipArchiveEntry, string> names = OfficeProvenanceZip.GetValidatedEntryNames(data, archive);
        long inspectedBytes = 0;
        long maximumInspectedBytes = limits.MaxExpandedContainerBytes;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            limits.CancellationToken.ThrowIfCancellationRequested();
            string name = names[entry];
            if (!OdfPackage.IsSignatureCandidatePath(name)) continue;
            if (OdfPackage.IsSignaturePath(name)) {
                signatures.Add(name);
                continue;
            }
            if (entry.Length > 1024 * 1024) {
                byte[] prefix = new byte[8];
                int read = 0;
                using (Stream prefixStream = entry.Open()) {
                    while (read < prefix.Length) {
                        int current = prefixStream.Read(prefix, read, prefix.Length - read);
                        if (current == 0) break;
                        read += current;
                    }
                }
                if (!OdfPackage.IsPngContent(prefix, read)) signatures.Add(name);
                continue;
            }
            if (entry.Length < 0 || entry.Length > maximumInspectedBytes - inspectedBytes) {
                throw OfficeProvenanceLimitException.Create("ODF signature classification exceeds the configured expanded-byte limit.");
            }
            inspectedBytes += entry.Length;
            using Stream content = entry.Open();
            byte[] bytes = OfficeProvenanceBinary.ReadBounded(content, 1024 * 1024, limits.CancellationToken);
            if (OdfPackage.IsSignatureEntry(name, bytes)) signatures.Add(name);
        }
        return signatures;
    }

    private static byte[] RemoveManifestEntries(
        byte[] data,
        OfficeProvenanceOptions limits,
        Func<string, bool> shouldRemove) {
        limits.CancellationToken.ThrowIfCancellationRequested();
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(data, limits, "ODF manifest");
        XDocument document;
        using (var stream = new MemoryStream(data, writable: false))
        using (XmlReader reader = XmlReader.Create(stream, OfficeProvenanceXml.CreateReaderSettings(limits))) {
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        limits.CancellationToken.ThrowIfCancellationRequested();
        XNamespace manifestNamespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        foreach (XElement entry in document.Descendants(manifestNamespace + "file-entry").ToArray()) {
            limits.CancellationToken.ThrowIfCancellationRequested();
            string? path = (string?)entry.Attribute(manifestNamespace + "full-path");
            if (path != null && shouldRemove(path)) entry.Remove();
        }
        using var output = new OfficeProvenanceBoundedMemoryStream(limits.MaxAssetBytes);
        using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
            Encoding = new System.Text.UTF8Encoding(false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration == null
        })) document.Save(writer);
        return output.ToArray();
    }
}
