using OfficeIMO.Provenance;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
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
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    /// <summary>Removes selected provenance from encoded ODF package bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] packageBytes,
        string fileName = "document.odt",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(packageBytes, fileName, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

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
            (_, manifest) => RemoveSignatureManifestEntries(manifest, limits),
            limits.MaxAssetBytes);
    }

    private static byte[] RemoveSignatureManifestEntries(byte[] data, OfficeProvenanceOptions limits) {
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(data, limits, "ODF manifest");
        XDocument document;
        using (var stream = new MemoryStream(data, writable: false))
        using (XmlReader reader = XmlReader.Create(stream, OfficeProvenanceXml.CreateReaderSettings(limits))) {
            document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        XNamespace manifestNamespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        foreach (XElement entry in document.Descendants(manifestNamespace + "file-entry").ToArray()) {
            string? path = (string?)entry.Attribute(manifestNamespace + "full-path");
            if (path != null && OdfPackage.IsSignaturePath(path)) entry.Remove();
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
