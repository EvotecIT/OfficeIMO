using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Provenance;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private const string SignatureOriginRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
    private const string SignaturePartRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

    /// <summary>Inspects C2PA and IPTC provenance in a saved VSDX package and its supported embedded images.</summary>
    public static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) =>
        OfficeProvenanceInspector.InspectFile(filePath, options);

    /// <summary>Removes selected provenance and atomically writes a VSDX package.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures);

    /// <summary>Removes selected provenance from encoded VSDX package bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] packageBytes,
        string fileName = "drawing.vsdx",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(packageBytes, fileName, options, StripPackageSignatures);

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceOptions limits) {
        using var stream = new MemoryStream(data.Length);
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        bool hadSignatures = false;
        using (Package package = Package.Open(stream, FileMode.Open, FileAccess.ReadWrite)) {
            PackageRelationship[] origins = package.GetRelationshipsByType(SignatureOriginRelationship).ToArray();
            foreach (PackageRelationship relationship in origins) {
                hadSignatures = true;
                if (relationship.TargetMode == TargetMode.External) {
                    package.DeleteRelationship(relationship.Id);
                    continue;
                }
                Uri originUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                if (package.PartExists(originUri)) {
                    PackagePart origin = package.GetPart(originUri);
                    foreach (PackageRelationship signature in origin.GetRelationshipsByType(SignaturePartRelationship).ToArray()) {
                        if (signature.TargetMode == TargetMode.External) continue;
                        Uri signatureUri = PackUriHelper.ResolvePartUri(origin.Uri, signature.TargetUri);
                        if (package.PartExists(signatureUri)) package.DeletePart(signatureUri);
                    }
                    package.DeletePart(originUri);
                }
                package.DeleteRelationship(relationship.Id);
            }

            Uri appPropertiesUri = PackUriHelper.CreatePartUri(new Uri("/docProps/app.xml", UriKind.Relative));
            if (package.PartExists(appPropertiesUri)) {
                PackagePart appProperties = package.GetPart(appPropertiesUri);
                XDocument? document = null;
                using (Stream input = appProperties.GetStream(FileMode.Open, FileAccess.Read)) {
                    if (!input.CanSeek || input.Length != 0) {
                        using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings {
                            DtdProcessing = DtdProcessing.Prohibit,
                            XmlResolver = null,
                            MaxCharactersInDocument = limits.MaxAssetBytes
                        });
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
}
