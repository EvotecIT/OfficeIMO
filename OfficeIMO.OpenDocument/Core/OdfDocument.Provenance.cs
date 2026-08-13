using OfficeIMO.Provenance;
using OfficeIMO.Core.Internal;
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
        using (Stream stream = File.OpenRead(fullPath)) data = OfficeProvenanceBinary.ReadBounded(stream, options.MaxAssetBytes);
        ValidatePackage(data, options);
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
        using (Stream stream = File.OpenRead(Path.GetFullPath(inputPath))) data = OfficeProvenanceBinary.ReadBounded(stream, options.Limits.MaxAssetBytes);
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
        ValidatePackage(packageBytes, options.Limits);
        bool hadManifest = OfficeProvenanceZip.HasEntry(packageBytes, path => path == ProvenanceManifestPath);
        OfficeProvenanceRemovalResult result = OfficeProvenancePackageMutation.Remove(
            packageBytes, fileName, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);
        byte[] output = result.ToArray();
        if (!hadManifest || OfficeProvenanceZip.HasEntry(output, path => path == ProvenanceManifestPath)) return result;
        OfficeProvenanceSignatureStripResult cleaned = OfficeProvenanceZip.RemoveEntries(
            output,
            _ => false,
            path => path == "META-INF/manifest.xml",
            (_, manifest) => RemoveManifestEntries(manifest, options.Limits, path => path == ProvenanceManifestPath),
            options.Limits.MaxAssetBytes);
        byte[] cleanedData = cleaned.Data;
        return new OfficeProvenanceRemovalResult(
            cleanedData,
            result.Before,
            result.After,
            result.Changes,
            wasReserialized: true,
            wereInvalidatedSignaturesRemoved: result.WereInvalidatedSignaturesRemoved);
    }

    private static readonly string[] SupportedMimetypes = {
        OdfMediaTypes.Text,
        OdfMediaTypes.Spreadsheet,
        OdfMediaTypes.Presentation
    };

    private static void ValidatePackage(byte[] data, OfficeProvenanceOptions options) =>
        OfficeProvenanceZip.ValidateMimetypeEntry(data, SupportedMimetypes, options.MaxContainerEntries);

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions _) =>
        OfficeProvenanceZip.HasEntry(data, OdfPackage.IsSignaturePath);

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceOptions limits) {
        return OfficeProvenanceZip.RemoveEntries(
            data,
            OdfPackage.IsSignaturePath,
            path => path == "META-INF/manifest.xml",
            (_, manifest) => RemoveManifestEntries(manifest, limits, OdfPackage.IsSignaturePath),
            limits.MaxAssetBytes);
    }

    private static byte[] RemoveManifestEntries(byte[] data, OfficeProvenanceOptions limits, Func<string, bool> shouldRemove) {
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(data, limits, "ODF manifest");
        XDocument document;
        using (var stream = new MemoryStream(data, writable: false))
        using (XmlReader reader = XmlReader.Create(stream, OfficeProvenanceXml.CreateReaderSettings(limits))) {
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        XNamespace manifestNamespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        foreach (XElement entry in document.Descendants(manifestNamespace + "file-entry").ToArray()) {
            string? path = (string?)entry.Attribute(manifestNamespace + "full-path");
            if (path != null && shouldRemove(path)) entry.Remove();
        }
        using var output = new MemoryStream();
        using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
            Encoding = new System.Text.UTF8Encoding(false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration == null
        })) document.Save(writer);
        return output.ToArray();
    }
}
