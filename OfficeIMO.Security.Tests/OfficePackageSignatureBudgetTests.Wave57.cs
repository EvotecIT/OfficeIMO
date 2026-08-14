using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureBudgetTestsWave57 {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConventionalApplicationMetadataIsDetectedWithoutAnExtendedPropertiesRelationship(bool includeRootRelationships) {
        var entries = new List<(string Name, byte[] Content)> {
            ("docProps/app.xml", Encoding.UTF8.GetBytes(
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><DigSig>present</DigSig></Properties>")),
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
                "</Types>"))
        };
        if (includeRootRelationships) {
            entries.Add(("_rels/.rels", Encoding.UTF8.GetBytes(
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                "</Relationships>")));
        }

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(CreatePackage(entries));

        Assert.True(info.HasApplicationSignatureMetadata);
        Assert.True(info.HasSignatures);
    }

    private static byte[] CreatePackage(IEnumerable<(string Name, byte[] Content)> entries) {
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
