using OfficeIMO.Provenance;

namespace OfficeIMO.Epub;

public sealed partial class EpubDocument {
    /// <summary>Inspects C2PA and IPTC provenance in an EPUB package and its supported embedded images.</summary>
    public static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) =>
        OfficeProvenanceInspector.InspectFile(filePath, options);

    /// <summary>Removes selected provenance and atomically writes an EPUB package.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures, HasPackageSignatures);

    /// <summary>Removes selected provenance from encoded EPUB package bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] packageBytes,
        string fileName = "publication.epub",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(packageBytes, fileName, options, StripPackageSignatures, HasPackageSignatures);

    private static bool HasPackageSignatures(byte[] data) => OfficeProvenanceZip.HasEntry(data, path =>
        path.Equals("META-INF/signatures.xml", StringComparison.OrdinalIgnoreCase));

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data) {
        return OfficeProvenanceZip.RemoveEntries(data, path =>
            path.Equals("META-INF/signatures.xml", StringComparison.OrdinalIgnoreCase));
    }
}
