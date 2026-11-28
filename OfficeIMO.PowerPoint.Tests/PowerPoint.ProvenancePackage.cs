using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

public sealed class PowerPointProvenancePackageTests {
    private const string Application = "/metadata/application.xml";
    private const string ApplicationType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    private const string ApplicationRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "extended-properties";
    private const string SignatureType = "application/vnd.openxmlformats-package.digital-signature-";
    private const string SignatureRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/";
    private const string Origin = "/custom-signatures/origin.sigs";
    private const string Signature = "/custom-signatures/certificate-owner.xml";
    private const string Certificate = "/custom-signatures/certificate.cer";
    private const string Manifest = "/META-INF/content_credential.c2pa";
    private static readonly XNamespace Properties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private static readonly XNamespace StrictProperties = "http://purl.oclc.org/ooxml/officeDocument/extendedProperties";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    [Theory]
    [InlineData(PresentationDocumentType.Presentation, ".pptx")]
    [InlineData(PresentationDocumentType.MacroEnabledPresentation, ".pptm")]
    [InlineData(PresentationDocumentType.AddIn, ".ppam")]
    [InlineData(PresentationDocumentType.Template, ".potx")]
    [InlineData(PresentationDocumentType.MacroEnabledTemplate, ".potm")]
    [InlineData(PresentationDocumentType.Slideshow, ".ppsx")]
    [InlineData(PresentationDocumentType.MacroEnabledSlideshow, ".ppsm")]
    public void ValidatesTheRelationshipOwnedMainPartSubtype(PresentationDocumentType type, string extension) {
        byte[] data = CreatePresentation(type);

        OfficeProvenanceRemovalResult result = PowerPointDocument.RemoveProvenance(data, "input" + extension);

        Assert.False(result.WasChanged);
        Assert.Equal(data, result.ToArray());
        string wrongExtension = extension == ".pptx" ? ".ppsx" : ".pptx";
        Assert.Throws<InvalidDataException>(() => PowerPointDocument.RemoveProvenance(data, "input" + wrongExtension));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SignatureCleanupRemovesOnlyTypedApplicationMetadataAndUnsharedSignatureParts(bool nativeOrigin) {
        byte[] data = CreateSignedPresentation(nativeOrigin, includeDirectSignature: true);

        Assert.Throws<InvalidOperationException>(() =>
            PowerPointDocument.RemoveProvenance(data, options: RemovalOptions(OfficeSignatureMutationPolicy.BlockSave)));

        OfficeProvenanceRemovalResult result = PowerPointDocument.RemoveProvenance(data,
            options: RemovalOptions(OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures));

        Assert.True(result.WereInvalidatedSignaturesRemoved);
        using var package = OfficeOpcPackage.Open(new MemoryStream(result.ToArray(), writable: false));
        Assert.False(package.PartExists(Manifest));
        Assert.False(package.PartExists(Origin));
        Assert.False(package.PartExists(Signature));
        Assert.False(package.PartExists(OfficeOpcPartNames.GetRelationshipsPartName(Origin)));
        Assert.False(package.PartExists(OfficeOpcPartNames.GetRelationshipsPartName(Signature)));
        Assert.Null(package.GetRelatedPartName(null, SignatureRelationship + "origin"));
        Assert.Equal(Application, package.GetRelatedPartName(null, ApplicationRelationship));
        if (nativeOrigin) {
            Assert.Equal(Certificate, package.GetRelatedPartName(Application, SignatureRelationship + "certificate"));
            using OfficePooledBytes certificate = package.ReadPart(Certificate);
            Assert.Equal(11, certificate.Array[0]);
        }
        XElement properties = ReadApplication(package);
        Assert.Empty(properties.Elements(Properties + "DigSig"));
        Assert.NotNull(properties.Element(StrictProperties + "DigSig"));
        Assert.NotNull(properties.Element(Properties + "extLst")?.Element(Properties + "DigSig"));
        Assert.NotNull(properties.Element(Mc + "AlternateContent")?.Element(Mc + "Fallback")?.Element(Properties + "DigSig"));
        Assert.Equal("Retained application value", (string?)properties.Element(Properties + "Application"));
    }

    [Fact]
    public void ApplicationSignaturesInUnknownMarkupDoNotBlockProvenanceRemoval() {
        byte[] data = CreateSignedPresentation(nativeOrigin: false, includeDirectSignature: false);

        OfficeProvenanceRemovalResult result = PowerPointDocument.RemoveProvenance(data,
            options: RemovalOptions(OfficeSignatureMutationPolicy.BlockSave));

        Assert.True(result.WasChanged);
        Assert.False(result.WereInvalidatedSignaturesRemoved);
        using var package = OfficeOpcPackage.Open(new MemoryStream(result.ToArray(), writable: false));
        Assert.Equal(2, ReadApplication(package).Descendants(Properties + "DigSig").Count());
    }

    [Fact]
    public void RejectsDanglingPartDuringApplicationSignatureCleanup() {
        byte[] data = EditPresentation(CreateSignedPresentation(nativeOrigin: false, includeDirectSignature: true), package =>
            package.AddRelationship("/ppt/presentation.xml", OfficeOpcPackage.SlideRelationshipType, "/ppt/slides/missing.xml"));

        Assert.Throws<InvalidDataException>(() => PowerPointDocument.RemoveProvenance(data,
            options: RemovalOptions(OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures)));
    }

    [Fact]
    public void RejectsConflictingIncomingTypesOnSharedSignatureParts() {
        byte[] data = EditPresentation(CreateSignedPresentation(nativeOrigin: true, includeDirectSignature: false), package =>
            package.AddRelationship(Application, "urn:other", Certificate));

        Assert.Throws<InvalidDataException>(() => PowerPointDocument.RemoveProvenance(data,
            options: RemovalOptions(OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures)));
    }

    [Fact]
    public void RejectsRelationshipsUnderAMalformedPackageRoot() {
        byte[] data = EditPresentation(CreatePresentation(PresentationDocumentType.Presentation), package => {
            const string path = "/_rels/.rels";
            using var root = new MemoryStream(package.ReadPartBytes(path), writable: false);
            XElement relationships = XElement.Load(root);
            relationships.Name = relationships.Name.Namespace + "Bogus";
            byte[] xml = Encoding.UTF8.GetBytes(relationships.ToString(SaveOptions.DisableFormatting));
            package.ReplacePart(path, xml, xml.Length);
        });

        Assert.Throws<InvalidDataException>(() => PowerPointDocument.InspectProvenance(data));
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void SignatureCleanupPreservesArchiveAndRetainedEntryMetadata() {
        byte[] package = CreateSignedPresentation(nativeOrigin: true, includeDirectSignature: true);
        using var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        stream.Position = 0;
        using (var input = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            input.Comment = "Presentation archive";
            ZipArchiveEntry entry = input.GetEntry(Application.TrimStart('/'))!;
            entry.Comment = "Application metadata";
            entry.ExternalAttributes = 0x20;
        }

        OfficeProvenanceRemovalResult result = PowerPointDocument.RemoveProvenance(stream.ToArray(),
            options: RemovalOptions(OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures));

        Assert.True(result.WereInvalidatedSignaturesRemoved);
        using var output = new ZipArchive(new MemoryStream(result.ToArray(), writable: false), ZipArchiveMode.Read);
        Assert.Equal("Presentation archive", output.Comment);
        ZipArchiveEntry retained = output.GetEntry(Application.TrimStart('/'))!;
        Assert.Equal("Application metadata", retained.Comment);
        Assert.Equal(0x20, retained.ExternalAttributes);
    }
#endif

    private static byte[] CreatePresentation(PresentationDocumentType type) {
        using var destination = new MemoryStream();
        using (var document = PresentationDocument.Create(destination, type)) {
            document.AddPresentationPart().Presentation = new P.Presentation();
        }
        return destination.ToArray();
    }

    private static byte[] CreateSignedPresentation(bool nativeOrigin, bool includeDirectSignature) {
        XElement root = new(Properties + "Properties",
            new XAttribute(XNamespace.Xmlns + "mc", Mc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "ep", Properties.NamespaceName),
            new XElement(Properties + "Application", "Retained application value"),
            new XElement(StrictProperties + "DigSig"),
            new XElement(Properties + "extLst", new XElement(Properties + "DigSig")),
            new XElement(Mc + "AlternateContent",
                new XElement(Mc + "Choice", new XAttribute("Requires", "ep")),
                new XElement(Mc + "Fallback", new XElement(Properties + "DigSig"))));
        if (includeDirectSignature) {
            root.AddFirst(new XElement(Properties + "DigSig"), new XElement(Properties + "DigSig"));
        }
        return EditPresentation(CreatePresentation(PresentationDocumentType.Presentation), package => {
            package.CreateBinaryPart(Application, ApplicationType, Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting)));
            package.AddRelationship(null, ApplicationRelationship, Application);
            package.CreateBinaryPart(Manifest, "application/c2pa", new byte[] { 1, 2, 3 });
            if (nativeOrigin) {
                package.CreateBinaryPart(Origin, SignatureType + "origin", Array.Empty<byte>());
                package.CreateBinaryPart(Signature, SignatureType + "xmlsignature+xml", Encoding.UTF8.GetBytes(
                    "<Signature xmlns='http://www.w3.org/2000/09/xmldsig#' />"));
                package.CreateBinaryPart(Certificate, SignatureType + "certificate", new byte[] { 11 });
                package.AddRelationship(null, SignatureRelationship + "origin", Origin);
                package.AddRelationship(Origin, SignatureRelationship + "signature", Signature);
                package.AddRelationship(Signature, SignatureRelationship + "certificate", Certificate);
                package.AddRelationship(Application, SignatureRelationship + "certificate", Certificate);
            }
        });
    }

    private static byte[] EditPresentation(byte[] data, Action<OfficeOpcEditablePackage> edit) {
        using var package = OfficeOpcEditablePackage.Open(new MemoryStream(data, writable: false));
        edit(package);
        using var output = new MemoryStream();
        package.Save(output);
        return output.ToArray();
    }

    private static XElement ReadApplication(OfficeOpcPackage package) {
        using OfficePooledBytes bytes = package.ReadPart(Application);
        using var source = new MemoryStream(bytes.Array, 0, bytes.Length, writable: false);
        return XElement.Load(source);
    }

    private static OfficeProvenanceRemovalOptions RemovalOptions(OfficeSignatureMutationPolicy signaturePolicy) => new() {
        RequireStructurallyValidCarrier = false,
        SignatureMutationPolicy = signaturePolicy
    };
}
