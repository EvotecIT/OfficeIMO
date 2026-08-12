using System.IO.Compression;
using OfficeIMO.Provenance;

namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    /// <summary>Inspects C2PA and IPTC provenance in an ODF package and its supported embedded images.</summary>
    public static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) =>
        OfficeProvenanceInspector.InspectFile(filePath, options);

    /// <summary>Removes selected provenance and atomically writes an ODF package.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures);

    /// <summary>Removes selected provenance from encoded ODF package bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] packageBytes,
        string fileName = "document.odt",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(packageBytes, fileName, options, StripPackageSignatures);

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data) {
        using var stream = new MemoryStream(data.Length);
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        bool hadSignatures = false;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry[] entries = archive.Entries.Where(item =>
                item.FullName.StartsWith("META-INF/", StringComparison.OrdinalIgnoreCase) &&
                item.FullName.EndsWith("signatures.xml", StringComparison.OrdinalIgnoreCase)).ToArray();
            foreach (ZipArchiveEntry entry in entries) {
                hadSignatures = true;
                entry.Delete();
            }
        }
        return new OfficeProvenanceSignatureStripResult(stream.ToArray(), hadSignatures);
    }
}
