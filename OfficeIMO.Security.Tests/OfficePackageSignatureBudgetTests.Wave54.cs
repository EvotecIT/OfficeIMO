using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureBudgetTestsWave54 {
    [Fact]
    public void UntransformedReferenceIsRejectedByTheRemainingDigestBudget() {
        byte[] payload = Enumerable.Repeat((byte)'x', 4096).ToArray();
        byte[] package = CreatePackage(
            ("payload.bin", payload),
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>" +
                "</Types>")));
        using var archive = new OfficePackageSignatureArchive(package);
        XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
        var reference = new XElement(ds + "Reference",
            new XAttribute("URI", "/payload.bin?ContentType=application%2Foctet-stream"),
            new XElement(ds + "DigestMethod", new XAttribute("Algorithm", XmlDigitalSignatureAlgorithms.Sha256)),
            new XElement(ds + "DigestValue", Convert.ToBase64String(SHA256.HashData(payload))));

        Assert.Throws<OfficePackageSignatureResourceLimitException>(() =>
            archive.VerifyReference(reference, maxPartBytes: payload.Length, maxDigestBytes: payload.Length - 1));
    }

    [Fact]
    public void RepeatedApplicationMetadataTargetsAreReadOnceWithinTheAggregateBudget() {
        const string relationships =
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
            "<Relationship Id=\"r2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/./app.xml\"/>" +
            "</Relationships>";
        const string properties =
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>OfficeIMO</Application></Properties>";
        byte[] relationshipBytes = Encoding.UTF8.GetBytes(relationships);
        byte[] propertyBytes = Encoding.UTF8.GetBytes(properties);
        byte[] package = CreatePackage(
            ("_rels/.rels", relationshipBytes),
            ("docProps/app.xml", propertyBytes),
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
                "</Types>")));

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(package, new OfficePackageSignatureInspectionOptions {
            VerifyDigests = false,
            MaxSignatureBytes = 4096,
            MaxTotalDigestBytes = relationshipBytes.Length + propertyBytes.Length
        });

        Assert.False(info.HasSignatures);
        Assert.DoesNotContain(info.Findings, finding => finding.Contains("aggregate limit", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ApplicationMetadataReadsRespectTheAggregateBudget() {
        const string relationships =
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
            "</Relationships>";
        byte[] relationshipBytes = Encoding.UTF8.GetBytes(relationships);
        byte[] propertyBytes = Encoding.UTF8.GetBytes(
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>" +
            new string('x', 1024) + "</Application></Properties>");
        byte[] package = CreatePackage(
            ("_rels/.rels", relationshipBytes),
            ("docProps/app.xml", propertyBytes),
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
                "</Types>")));

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(package, new OfficePackageSignatureInspectionOptions {
            VerifyDigests = false,
            MaxSignatureBytes = 4096,
            MaxTotalDigestBytes = relationshipBytes.Length + propertyBytes.Length - 1
        });

        Assert.False(info.SignatureDiscoveryComplete);
        Assert.Contains(info.Findings, finding => finding.Contains("aggregate limit", StringComparison.OrdinalIgnoreCase));
    }

    private static byte[] CreatePackage(params (string Name, byte[] Content)[] entries) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string name, byte[] content) in entries) {
                ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
                using Stream target = entry.Open();
                target.Write(content, 0, content.Length);
            }
        }
        return output.ToArray();
    }
}
