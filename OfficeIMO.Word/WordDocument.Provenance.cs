using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Provenance;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Validates and inspects encoded Word package bytes without accessing the filesystem.</summary>
    public static OfficeProvenanceReport InspectProvenance(byte[] data, string fileName = "document.docx", OfficeProvenanceOptions? options = null) =>
        OfficeProvenancePackageMutation.Inspect(data, fileName, options, ValidatePackage);

    /// <summary>Inspects C2PA and IPTC provenance in a saved Open XML document and its supported embedded images.</summary>
    public static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) =>
        OfficeProvenancePackageMutation.InspectFile(filePath, options, ValidatePackage);

    /// <summary>Removes selected provenance and atomically writes an Open XML document.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    /// <summary>Removes selected provenance from encoded Open XML document bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] documentBytes,
        string fileName = "document.docx",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(documentBytes, fileName, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    private static void ValidatePackage(byte[] data, OfficeProvenanceOptions _) {
        OfficeProvenanceZip.ValidateForOwningPackageMutation(data, _);
        using var stream = new MemoryStream(data, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
        if (_.RequireStandardOpenXmlDocument && document.DocumentType != DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
            throw new InvalidDataException("The memory-only workflow requires a DOCX document, not a macro-enabled document or template.");
        if (document.MainDocumentPart == null || !IsSupportedMainPartContentType(document.MainDocumentPart.ContentType)) {
            throw new InvalidDataException("The package is not a Word document.");
        }
    }

    private static bool IsSupportedMainPartContentType(string contentType) => new[] {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"
    }.Contains(contentType, StringComparer.OrdinalIgnoreCase);

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) =>
        OfficeProvenanceZip.HasNativePackageSignature(data, options) ||
        HasRelationshipOwnedApplicationSignatureMetadata(data, options.Limits);

    private static bool HasRelationshipOwnedApplicationSignatureMetadata(byte[] data, OfficeProvenanceOptions limits) {
        using var stream = new MemoryStream(data, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
        ExtendedFilePropertiesPart? applicationProperties = document.ExtendedFilePropertiesPart;
        ValidateApplicationProperties(applicationProperties, limits);
        return applicationProperties?.Properties?.DigitalSignature != null;
    }

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) {
        OfficeProvenanceOptions limits = options.Limits;
        using var stream = new OfficeProvenanceBoundedMemoryStream(options.EffectiveMaxIntermediateBytes, data.Length);
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        bool hadSignatures;
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, true)) {
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
