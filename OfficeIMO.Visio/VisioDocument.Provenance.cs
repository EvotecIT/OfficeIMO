using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Provenance;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private const string SignatureOriginRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
    private const string SignaturePartRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
    private const string DocumentRelationship = "http://schemas.microsoft.com/visio/2010/relationships/document";
    private const string SignatureOriginContentType = "application/vnd.openxmlformats-package.digital-signature-origin";
    private const string SignaturePartContentType = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";

    /// <summary>Inspects C2PA and IPTC provenance in a saved VSDX package and its supported embedded images.</summary>
    public static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) =>
        OfficeProvenancePackageMutation.InspectFile(filePath, options, ValidatePackage);

    /// <summary>Removes selected provenance and atomically writes a VSDX package.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    /// <summary>Removes selected provenance from encoded VSDX package bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] packageBytes,
        string fileName = "drawing.vsdx",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(packageBytes, fileName, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    private static void ValidatePackage(byte[] data, OfficeProvenanceOptions options) {
        OfficeProvenanceZip.ValidateForOwningPackageMutation(data, options);
        using var stream = new MemoryStream(data, writable: false);
        using Package package = Package.Open(stream, FileMode.Open, FileAccess.Read);
        PackageRelationship[] relationships = package.GetRelationshipsByType(DocumentRelationship).ToArray();
        if (relationships.Length != 1 || relationships[0].TargetMode != TargetMode.Internal) {
            throw new InvalidDataException("The package is not a supported Visio document.");
        }
        Uri documentUri = PackUriHelper.ResolvePartUri(relationships[0].SourceUri, relationships[0].TargetUri);
        if (!package.PartExists(documentUri) || !VisioPackageFormat.TryFromContentType(package.GetPart(documentUri).ContentType, out _)) {
            throw new InvalidDataException("The package is not a supported Visio document.");
        }
    }

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) =>
        OfficeProvenanceZip.HasPackageSignature(data, options) ||
        OfficeProvenanceZip.HasApplicationSignatureMetadata(data, options.Limits);

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceOptions limits) {
        using var stream = new MemoryStream(data.Length);
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        bool hadSignatures = false;
        using (Package package = Package.Open(stream, FileMode.Open, FileAccess.ReadWrite)) {
            PackageRelationship[] origins = package.GetRelationshipsByType(SignatureOriginRelationship).ToArray();
            var originParts = new List<PackagePart>();
            var signatureParts = new List<PackagePart>();
            foreach (PackageRelationship relationship in origins) {
                hadSignatures = true;
                if (relationship.TargetMode == TargetMode.External) continue;
                Uri originUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                if (package.PartExists(originUri)) {
                    PackagePart origin = package.GetPart(originUri);
                    if (!string.Equals(origin.ContentType, SignatureOriginContentType, StringComparison.OrdinalIgnoreCase)) {
                        throw new InvalidDataException("A Visio signature-origin relationship targets a part with an unexpected content type.");
                    }
                    originParts.Add(origin);
                    foreach (PackageRelationship signature in origin.GetRelationshipsByType(SignaturePartRelationship).ToArray()) {
                        if (signature.TargetMode == TargetMode.External) continue;
                        Uri signatureUri = PackUriHelper.ResolvePartUri(origin.Uri, signature.TargetUri);
                        if (!package.PartExists(signatureUri)) continue;
                        PackagePart signaturePart = package.GetPart(signatureUri);
                        if (!string.Equals(signaturePart.ContentType, SignaturePartContentType, StringComparison.OrdinalIgnoreCase)) {
                            throw new InvalidDataException("A Visio signature relationship targets a part with an unexpected content type.");
                        }
                        signatureParts.Add(signaturePart);
                    }
                }
            }
            foreach (PackagePart signaturePart in signatureParts.Distinct()) package.DeletePart(signaturePart.Uri);
            foreach (PackagePart originPart in originParts.Distinct()) package.DeletePart(originPart.Uri);
            foreach (PackageRelationship relationship in origins) {
                package.DeleteRelationship(relationship.Id);
            }

            Uri appPropertiesUri = PackUriHelper.CreatePartUri(new Uri("/docProps/app.xml", UriKind.Relative));
            if (package.PartExists(appPropertiesUri)) {
                PackagePart appProperties = package.GetPart(appPropertiesUri);
                XDocument? document = null;
                using (Stream input = appProperties.GetStream(FileMode.Open, FileAccess.Read)) {
                    byte[] xml = ReadBoundedXml(input, limits.MaxAssetBytes);
                    if (xml.Length != 0) {
                        OfficeProvenanceXml.ValidateMaterializedNodeBudget(xml, limits, "Visio app metadata");
                        using var xmlInput = new MemoryStream(xml, writable: false);
                        using XmlReader reader = XmlReader.Create(xmlInput, OfficeProvenanceXml.CreateReaderSettings(limits));
                        document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
                    }
                }
                XNamespace properties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
                XElement[] digitalSignatures = document?.Root?.Elements(properties + "DigSig").ToArray() ?? Array.Empty<XElement>();
                if (digitalSignatures.Length != 0) {
                    hadSignatures = true;
                    foreach (XElement element in digitalSignatures) element.Remove();
                    using Stream output = appProperties.GetStream(FileMode.Create, FileAccess.Write);
                    document!.Save(output, SaveOptions.DisableFormatting);
                }
            }
        }
        return new OfficeProvenanceSignatureStripResult(stream.ToArray(), hadSignatures);
    }

    private static byte[] ReadBoundedXml(Stream input, long maximumBytes) {
        if (input.CanSeek && input.Length > maximumBytes) {
            throw new InvalidDataException("Visio app metadata exceeds the configured asset limit.");
        }
        using var output = new MemoryStream();
        byte[] buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = input.Read(buffer, 0, buffer.Length)) > 0) {
            total += read;
            if (total > maximumBytes) throw new InvalidDataException("Visio app metadata exceeds the configured asset limit.");
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

}
