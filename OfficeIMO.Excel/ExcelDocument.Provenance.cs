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
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    /// <summary>Removes selected provenance from encoded Open XML workbook bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] workbookBytes,
        string fileName = "workbook.xlsx",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(workbookBytes, fileName, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    private static void ValidatePackage(byte[] data, OfficeProvenanceOptions _) {
        OfficeProvenanceZip.ValidateForOwningPackageMutation(data, _);
        using var stream = new MemoryStream(data, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
        if (document.WorkbookPart == null || document.WorkbookPart.ContentType is not (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" or
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml" or
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml" or
            "application/vnd.ms-excel.template.macroEnabled.main+xml" or
            "application/vnd.ms-excel.addin.macroEnabled.main+xml")) {
            throw new InvalidDataException("The package is not an Excel workbook.");
        }
    }

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) =>
        OfficeProvenanceZip.HasPackageSignature(data, options) ||
        OfficeProvenanceZip.HasApplicationSignatureMetadata(data, options.Limits);

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceOptions limits) {
        using var stream = new MemoryStream(data.Length);
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        bool hadSignatures;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true)) {
            DigitalSignatureOriginPart? origin = document.DigitalSignatureOriginPart;
            ExtendedFilePropertiesPart? applicationProperties = document.ExtendedFilePropertiesPart;
            ValidateApplicationProperties(applicationProperties, limits);
            bool hasApplicationMetadata = applicationProperties?.Properties?.DigitalSignature != null;
            hadSignatures = origin != null || hasApplicationMetadata;
            if (origin != null) document.DeletePart(origin);
            if (hasApplicationMetadata) {
                document.ExtendedFilePropertiesPart!.Properties!.DigitalSignature = null;
                document.ExtendedFilePropertiesPart.Properties.Save();
            }
        }
        return new OfficeProvenanceSignatureStripResult(stream.ToArray(), hadSignatures);
    }

    private static void ValidateApplicationProperties(ExtendedFilePropertiesPart? part, OfficeProvenanceOptions limits) {
        if (part == null) return;
        using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(input, limits, "Open XML application metadata");
    }
}
