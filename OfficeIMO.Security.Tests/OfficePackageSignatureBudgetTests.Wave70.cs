using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureBudgetTestsWave70 {
    [Fact]
    public void ApplicationMetadataBudgetFailureDoesNotReportExcessSignatureParts() {
        byte[] properties = Encoding.UTF8.GetBytes(
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>OfficeIMO</Application></Properties>");
        byte[] package = CreatePackage(("docProps/app.xml", properties), ("[Content_Types].xml", Encoding.UTF8.GetBytes(
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
            "</Types>")));

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(package, new OfficePackageSignatureInspectionOptions {
            MaxTotalDigestBytes = properties.Length - 1
        });

        Assert.False(info.SignatureDiscoveryComplete);
        Assert.DoesNotContain(info.Findings, finding => finding.Contains("more XML signature parts", StringComparison.Ordinal));
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
