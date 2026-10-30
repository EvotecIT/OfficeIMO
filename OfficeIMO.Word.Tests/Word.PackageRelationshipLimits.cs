using System.IO.Compression;
using OfficeIMO;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPackageRelationshipLimitsTests {
    [Fact]
    public void PackageInspectorStopsBeforeUnboundedRelationshipTraversal() {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true)) {
            Write(archive, "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
            Write(archive, "_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                string.Concat(Enumerable.Range(1, 3).Select(index =>
                    $"<Relationship Id=\"rId{index}\" Type=\"http://example.org/link\" Target=\"https://example.org/{index}\" TargetMode=\"External\"/>")) +
                "</Relationships>");
        }
        byte[] package = buffer.ToArray();

        OfficePackageSecurityException exceeded = Assert.Throws<OfficePackageSecurityException>(() =>
            OfficePackageSecurityInspector.Validate(package, new OfficePackageSecurityOptions { MaxRelationships = 2 }));
        Assert.Equal(OfficePackageSecurityRule.RelationshipCount, exceeded.Rule);
        Assert.True(OfficePackageSecurityInspector.Validate(package,
            new OfficePackageSecurityOptions { MaxRelationships = 3 }).IsValid);
    }

    private static void Write(ZipArchive archive, string name, string xml) {
        using StreamWriter writer = new(archive.CreateEntry(name).Open());
        writer.Write(xml);
    }
}
