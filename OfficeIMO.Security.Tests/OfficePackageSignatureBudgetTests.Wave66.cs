using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureBudgetTestsWave66 {
    [Fact]
    public void LaterExtendedPropertiesLimitPreservesConventionalSignatureMetadataEvidence() {
        byte[] conventional = Encoding.UTF8.GetBytes(
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><DigSig>present</DigSig></Properties>");
        byte[] relationships = Encoding.UTF8.GetBytes(
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/extra.xml\"/>" +
            "</Relationships>");
        byte[] extra = Encoding.UTF8.GetBytes(
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>" +
            new string('x', 1024) + "</Application></Properties>");
        byte[] package = CreatePackage(
            ("docProps/app.xml", conventional),
            ("docProps/extra.xml", extra),
            ("_rels/.rels", relationships),
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
                "<Override PartName=\"/docProps/extra.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
                "</Types>")));

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(package, new OfficePackageSignatureInspectionOptions {
            VerifyDigests = false,
            MaxSignatureBytes = 4096,
            MaxTotalDigestBytes = conventional.Length + relationships.Length
        });

        Assert.True(info.HasApplicationSignatureMetadata);
        Assert.False(info.SignatureDiscoveryComplete);
        Assert.Contains(info.Findings, finding => finding.Contains("resource limit", StringComparison.OrdinalIgnoreCase));
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
