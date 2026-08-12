using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Provenance;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Inspects C2PA and IPTC provenance in a saved Open XML workbook and its supported embedded images.</summary>
    public static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) =>
        OfficeProvenanceInspector.InspectFile(filePath, options);

    /// <summary>Removes selected provenance and atomically writes an Open XML workbook.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures);

    /// <summary>Removes selected provenance from encoded Open XML workbook bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] workbookBytes,
        string fileName = "workbook.xlsx",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(workbookBytes, fileName, options, StripPackageSignatures);

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data) {
        using var stream = new MemoryStream(data.Length);
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        bool hadSignatures;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true)) {
            DigitalSignatureOriginPart? origin = document.DigitalSignatureOriginPart;
            bool hasApplicationMetadata = document.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null;
            hadSignatures = origin != null || hasApplicationMetadata;
            if (origin != null) document.DeletePart(origin);
            if (hasApplicationMetadata) {
                document.ExtendedFilePropertiesPart!.Properties!.DigitalSignature = null;
                document.ExtendedFilePropertiesPart.Properties.Save();
            }
        }
        return new OfficeProvenanceSignatureStripResult(stream.ToArray(), hadSignatures);
    }
}
