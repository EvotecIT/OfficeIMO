using System.IO.Compression;
using System.IO;
using System.Security.Cryptography;
using System.Xml.Linq;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureBudgetTestsWave47 {
    [Fact]
    public void DigestWorkIsReservedOnlyForReferencesThatReachHashing() {
        byte[] package = CreatePackage();
        using var archive = new OfficePackageSignatureArchive(package);
        XElement oversized = CreateReference("/large.bin", new byte[8]);
        XElement valid = CreateReference("/small.bin", new byte[] { 1, 2, 3 });

        OfficePackageDigestResult skipped = archive.VerifyReference(oversized, maxPartBytes: 4, maxDigestBytes: 3);
        OfficePackageDigestResult hashed = archive.VerifyReference(valid, maxPartBytes: 4, maxDigestBytes: 3);

        Assert.Equal(OfficePackageSignatureValidationState.Unsupported, skipped.Status);
        Assert.Equal(0, skipped.DigestWorkBytes);
        Assert.Equal(OfficePackageSignatureValidationState.Passed, hashed.Status);
        Assert.Equal(3, hashed.DigestWorkBytes);
    }

    private static XElement CreateReference(string path, byte[] content) {
        XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
        return new XElement(ds + "Reference",
            new XAttribute("URI", path + "?ContentType=application%2Foctet-stream"),
            new XElement(ds + "DigestMethod", new XAttribute("Algorithm", XmlDigitalSignatureAlgorithms.Sha256)),
            new XElement(ds + "DigestValue", Convert.ToBase64String(SHA256.HashData(content))));
    }

    private static byte[] CreatePackage() {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>" +
                "</Types>");
            WriteEntry(archive, "large.bin", new byte[8]);
            WriteEntry(archive, "small.bin", new byte[] { 1, 2, 3 });
        }
        return output.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, string content) =>
        WriteEntry(archive, name, Encoding.UTF8.GetBytes(content));

    private static void WriteEntry(ZipArchive archive, string name, byte[] content) {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
        using Stream target = entry.Open();
        target.Write(content, 0, content.Length);
    }
}
