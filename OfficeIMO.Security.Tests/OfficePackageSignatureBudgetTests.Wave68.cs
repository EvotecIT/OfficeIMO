using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureBudgetTestsWave68 {
    [Fact]
    public void DigestLimitFailureReportsAlreadyConsumedTransformInput() {
        byte[] payload = Encoding.UTF8.GetBytes("<root><value>" + new string('x', 100) + "</value></root>");
        byte[] package = CreatePackage(payload);
        using var archive = new OfficePackageSignatureArchive(package, securityProvider: new OfficeSecurityProvider());
        XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
        var reference = new XElement(ds + "Reference",
            new XAttribute("URI", "/payload.xml?ContentType=" + Uri.EscapeDataString("application/xml")),
            new XElement(ds + "Transforms",
                new XElement(ds + "Transform", new XAttribute("Algorithm", OfficePackageSignatureArchive.CanonicalXmlAlgorithm))),
            new XElement(ds + "DigestMethod", new XAttribute("Algorithm", XmlDigitalSignatureAlgorithms.Sha256)),
            new XElement(ds + "DigestValue", Convert.ToBase64String(SHA256.HashData(payload))));

        OfficePackageSignatureResourceLimitException exception = Assert.Throws<OfficePackageSignatureResourceLimitException>(
            () => archive.VerifyReference(reference, maxPartBytes: 4096, maxDigestBytes: payload.Length + 1));

        Assert.Equal(payload.Length, exception.ConsumedBytes);
    }

    private static byte[] CreatePackage(byte[] payload) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "[Content_Types].xml", Encoding.UTF8.GetBytes(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Override PartName=\"/payload.xml\" ContentType=\"application/xml\"/>" +
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
