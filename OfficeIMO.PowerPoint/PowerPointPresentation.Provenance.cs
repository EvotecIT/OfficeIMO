using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Provenance;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointPresentation {
    /// <summary>Inspects C2PA and IPTC provenance in a saved Open XML presentation and its supported embedded images.</summary>
    public static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) =>
        OfficeProvenancePackageMutation.InspectFile(filePath, options, ValidatePackage);

    /// <summary>Removes selected provenance and atomically writes an Open XML presentation.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    /// <summary>Removes selected provenance from encoded Open XML presentation bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] presentationBytes,
        string fileName = "presentation.pptx",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(presentationBytes, fileName, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    private static void ValidatePackage(byte[] data, OfficeProvenanceOptions _) {
        OfficeProvenanceZip.ValidateForOwningPackageMutation(data, _);
        using var stream = new MemoryStream(data, writable: false);
        using PresentationDocument document = PresentationDocument.Open(stream, false);
        if (document.PresentationPart == null || !IsSupportedPresentationContentType(document.PresentationPart.ContentType)) {
            throw new InvalidDataException("The package is not a PowerPoint presentation.");
        }
    }

    private static bool IsSupportedPresentationContentType(string contentType) => new[] {
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
        "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml",
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
        "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
        "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
        "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml"
    }.Contains(contentType, StringComparer.OrdinalIgnoreCase);

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) =>
        OfficeProvenanceZip.HasNativePackageSignature(data, options) ||
        HasRelationshipOwnedApplicationSignatureMetadata(data, options.Limits);

    private static bool HasRelationshipOwnedApplicationSignatureMetadata(byte[] data, OfficeProvenanceOptions limits) {
        using var stream = new MemoryStream(data, writable: false);
        using PresentationDocument document = PresentationDocument.Open(stream, false);
        ExtendedFilePropertiesPart? applicationProperties = document.ExtendedFilePropertiesPart;
        ValidateApplicationProperties(applicationProperties, limits);
        return applicationProperties?.Properties?.DigitalSignature != null;
    }

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) {
        OfficeProvenanceOptions limits = options.Limits;
        using var stream = new OfficeProvenanceBoundedMemoryStream(options.EffectiveMaxOutputBytes, data.Length);
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        bool hadSignatures;
        using (PresentationDocument document = PresentationDocument.Open(stream, true)) {
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
