using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureTransformTestsWave53 {
    [Fact]
    public void RelationshipTransformRejectsAnOrdinaryXmlPart() {
        byte[] package = CreatePackage(
            "application/xml",
            Encoding.UTF8.GetBytes("<document><value>mutable</value></document>"));
        using var archive = new OfficePackageSignatureArchive(package);
        XElement reference = CreateRelationshipTransformReference("application/xml");

        Assert.Throws<InvalidDataException>(() => archive.ComputeDigestValue(reference, maxPartBytes: 4096));
        OfficePackageDigestResult result = archive.VerifyReference(reference, maxPartBytes: 4096);

        Assert.Equal(OfficePackageSignatureValidationState.Failed, result.Status);
        Assert.Contains("relationships part", result.Detail, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RelationshipTransformRejectsAMalformedRelationshipsDocument() {
        byte[] package = CreatePackage(
            OfficePackageSignatureArchive.RelationshipsContentType,
            Encoding.UTF8.GetBytes("<document xmlns=\"urn:not-opc\"><Relationship Id=\"rId1\"/></document>"));
        using var archive = new OfficePackageSignatureArchive(package);
        XElement reference = CreateRelationshipTransformReference(OfficePackageSignatureArchive.RelationshipsContentType);

        Assert.Throws<InvalidDataException>(() => archive.ComputeDigestValue(reference, maxPartBytes: 4096));
        OfficePackageDigestResult result = archive.VerifyReference(reference, maxPartBytes: 4096);

        Assert.Equal(OfficePackageSignatureValidationState.Failed, result.Status);
        Assert.Contains("Relationships document", result.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void RelationshipTransformRejectsMalformedRelationshipEntries() {
        const string relationships =
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\"/>" +
            "</Relationships>";
        byte[] package = CreatePackage(
            OfficePackageSignatureArchive.RelationshipsContentType,
            Encoding.UTF8.GetBytes(relationships));
        using var archive = new OfficePackageSignatureArchive(package);
        XElement reference = CreateRelationshipTransformReference(OfficePackageSignatureArchive.RelationshipsContentType);

        Assert.Throws<InvalidDataException>(() => archive.ComputeDigestValue(reference, maxPartBytes: 4096));
        OfficePackageDigestResult result = archive.VerifyReference(reference, maxPartBytes: 4096);

        Assert.Equal(OfficePackageSignatureValidationState.Failed, result.Status);
        Assert.Contains("Relationships document", result.Detail, StringComparison.Ordinal);
    }

    private static XElement CreateRelationshipTransformReference(string contentType) {
        XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
        XNamespace opc = "http://schemas.openxmlformats.org/package/2006/digital-signature";
        return new XElement(ds + "Reference",
            new XAttribute("URI", "/payload.xml?ContentType=" + Uri.EscapeDataString(contentType)),
            new XElement(ds + "Transforms",
                new XElement(ds + "Transform",
                    new XAttribute("Algorithm", OfficePackageSignatureArchive.RelationshipTransformAlgorithm),
                    new XElement(opc + "RelationshipReference", new XAttribute("SourceId", "rId1"))),
                new XElement(ds + "Transform", new XAttribute("Algorithm", OfficePackageSignatureArchive.CanonicalXmlAlgorithm))),
            new XElement(ds + "DigestMethod", new XAttribute("Algorithm", XmlDigitalSignatureAlgorithms.Sha256)),
            new XElement(ds + "DigestValue", Convert.ToBase64String(SHA256.HashData(Array.Empty<byte>()))));
    }

    private static byte[] CreatePackage(string contentType, byte[] payload) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "[Content_Types].xml", Encoding.UTF8.GetBytes(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Override PartName=\"/payload.xml\" ContentType=\"" + contentType + "\"/>" +
                "</Types>"));
            WriteEntry(archive, "payload.xml", payload);
        }
        return output.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, byte[] content) {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
        using Stream target = entry.Open();
        target.Write(content, 0, content.Length);
    }
}
