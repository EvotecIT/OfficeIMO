using OfficeIMO.Provenance;

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
            inputPath, outputPath, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    /// <summary>Removes selected provenance from encoded EPUB package bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] packageBytes,
        string fileName = "publication.epub",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(
            packageBytes, fileName, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    private static void ValidatePackage(byte[] data, OfficeProvenanceOptions options) =>
        OfficeProvenanceZip.ValidateMimetypeEntry(data, "application/epub+zip", options.MaxContainerEntries);

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions _) => OfficeProvenanceZip.HasEntry(data, path =>
        path.Equals("META-INF/signatures.xml", StringComparison.Ordinal));

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceOptions _) {
        return OfficeProvenanceZip.RemoveEntries(data, path =>
            path.Equals("META-INF/signatures.xml", StringComparison.Ordinal));
    }
}
