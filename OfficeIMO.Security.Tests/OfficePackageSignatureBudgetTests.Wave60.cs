using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureBudgetTestsWave60 {
    [Fact]
    public void InspectionRetainsTheBoundedSignatureBytesUsedByValidation() {
        const string contentTypes =
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Override PartName=\"/_xmlsignatures/sig1.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "</Types>";
        const string signature =
            "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
            "<SignedInfo><Reference URI=\"/document.xml\"><DigestMethod Algorithm=\"sha256\"/><DigestValue>AA==</DigestValue></Reference></SignedInfo>" +
            "</Signature>";
        byte[] signatureBytes = Encoding.UTF8.GetBytes(signature);
        byte[] package = CreatePackage(
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(contentTypes)),
            ("_xmlsignatures/sig1.xml", signatureBytes),
            ("document.xml", Encoding.UTF8.GetBytes("<document/>")));

        OfficePackageSignaturePartInfo part = Assert.Single(
            OfficePackageSignatureService.Inspect(package).SignatureParts);

        Assert.Equal(signatureBytes, part.SignatureBytes);
    }

    private static byte[] CreatePackage(params (string Name, byte[] Content)[] entries) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string name, byte[] content) in entries) {
                ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
                using Stream stream = entry.Open();
                stream.Write(content, 0, content.Length);
            }
        }
        return output.ToArray();
    }
}
