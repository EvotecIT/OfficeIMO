using System.IO.Compression;
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
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures);

    /// <summary>Removes selected provenance from encoded EPUB package bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] packageBytes,
        string fileName = "publication.epub",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(packageBytes, fileName, options, StripPackageSignatures);

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data) {
        using var stream = new MemoryStream(data.Length);
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        bool hadSignatures = false;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry[] entries = archive.Entries.Where(item =>
                item.FullName.Equals("META-INF/signatures.xml", StringComparison.OrdinalIgnoreCase)).ToArray();
            foreach (ZipArchiveEntry entry in entries) {
                hadSignatures = true;
                entry.Delete();
            }
        }
        return new OfficeProvenanceSignatureStripResult(stream.ToArray(), hadSignatures);
    }
}
