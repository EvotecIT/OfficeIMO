using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureBudgetTestsWave48 {
    [Fact]
    public void RelationshipTransformInputIsChargedBeforeLaterUnsupportedTransforms() {
        byte[] relationships = Encoding.UTF8.GetBytes(
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"urn:test\" Target=\"part.bin\"/>" +
            "</Relationships>");
        byte[] package = CreatePackage(("_rels/.rels", relationships));
        using var archive = new OfficePackageSignatureArchive(package);
        XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
        XNamespace opc = "http://schemas.openxmlformats.org/package/2006/digital-signature";
        var reference = new XElement(ds + "Reference",
            new XAttribute("URI", "/_rels/.rels?ContentType=application%2Fvnd.openxmlformats-package.relationships%2Bxml"),
            new XElement(ds + "Transforms",
                new XElement(ds + "Transform",
                    new XAttribute("Algorithm", OfficePackageSignatureArchive.RelationshipTransformAlgorithm),
                    new XElement(opc + "RelationshipReference", new XAttribute("SourceId", "rId1"))),
                new XElement(ds + "Transform", new XAttribute("Algorithm", "urn:unsupported"))),
            new XElement(ds + "DigestMethod", new XAttribute("Algorithm", XmlDigitalSignatureAlgorithms.Sha256)),
            new XElement(ds + "DigestValue", Convert.ToBase64String(SHA256.HashData(Array.Empty<byte>()))));

        Assert.Throws<OfficePackageSignatureResourceLimitException>(() =>
            archive.VerifyReference(reference, maxPartBytes: 4096, maxDigestBytes: relationships.Length - 1));
    }

    [Fact]
    public void MalformedOriginTargetsRemainSignatureEvidence() {
        byte[] relationships = Encoding.UTF8.GetBytes(
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rSig\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"https://outside.example/origin.sigs\"/>" +
            "</Relationships>");
        byte[] package = CreatePackage(("_rels/.rels", relationships));

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(
            package,
            new OfficePackageSignatureInspectionOptions { VerifyDigests = false });

        Assert.Equal(1, info.OriginRelationshipCount);
        Assert.True(info.HasSignatures);
        Assert.Contains(info.Findings, finding => finding.Contains("target is invalid", StringComparison.OrdinalIgnoreCase));
    }

    private static byte[] CreatePackage(params (string Name, byte[] Content)[] entries) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "[Content_Types].xml", Encoding.UTF8.GetBytes(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "</Types>"));
            foreach ((string name, byte[] content) in entries) WriteEntry(archive, name, content);
        }
        return output.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, byte[] content) {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
        using Stream target = entry.Open();
        target.Write(content, 0, content.Length);
    }
}
