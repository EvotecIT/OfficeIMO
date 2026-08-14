using System.IO.Compression;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void MalformedApplicationMetadataTargetDoesNotHideLaterSignatureEvidence() {
        const string relationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteZipEntry(archive, "[Content_Types].xml",
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
                    "</Types>");
                WriteZipEntry(archive, "_rels/.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"bad\" Type=\"" + relationshipType + "\" Target=\"https://example.invalid/app.xml\"/>" +
                    "<Relationship Id=\"valid\" Type=\"" + relationshipType + "\" Target=\"docProps/app.xml\"/>" +
                    "</Relationships>");
                WriteZipEntry(archive, "docProps/app.xml",
                    "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><DigSig>present</DigSig></Properties>");
            }
            package = output.ToArray();
        }

        OfficePackageSignatureInfo inspection = OfficePackageSignatureService.Inspect(package);

        Assert.True(inspection.HasApplicationSignatureMetadata);
        Assert.True(inspection.HasSignatures);
        Assert.Contains(inspection.Findings, finding =>
            finding.Contains("leaves the package namespace", StringComparison.OrdinalIgnoreCase));
    }
}
