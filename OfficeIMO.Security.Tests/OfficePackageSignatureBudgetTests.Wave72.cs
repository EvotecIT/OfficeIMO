using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureBudgetTestsWave72 {
    [Fact]
    public void MalformedRootRelationshipsMakeApplicationMetadataDiscoveryIncomplete() {
        byte[] package = CreatePackage(
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(
                "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>" +
                "<Override PartName='/custom/app.xml' ContentType='application/vnd.openxmlformats-officedocument.extended-properties+xml'/>" +
                "</Types>")),
            ("_rels/.rels", Encoding.UTF8.GetBytes("<Relationships")),
            ("custom/app.xml", Encoding.UTF8.GetBytes(
                "<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'><DigSig/></Properties>")));

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(package);

        Assert.False(info.SignatureDiscoveryComplete);
        Assert.Contains(info.Findings, finding => finding.Contains("could not be parsed", StringComparison.OrdinalIgnoreCase));
    }

    private static byte[] CreatePackage(params (string Name, byte[] Content)[] entries) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string name, byte[] content) in entries) {
                ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
                using Stream target = entry.Open();
                target.Write(content, 0, content.Length);
            }
        }
        return output.ToArray();
    }
}
