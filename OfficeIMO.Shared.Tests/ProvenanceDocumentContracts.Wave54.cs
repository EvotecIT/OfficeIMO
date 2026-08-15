using System.IO.Compression;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlStringRemovalRewritesOnlyAnExactCharsetParameter() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<html><head><meta http-equiv=\"content-type\" content=\"text/html; xcharset=windows-1252\">" +
            "<script type=\"application/c2pa\">" + manifest + "</script></head><body></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.True(result.WasChanged);
        Assert.Contains("xcharset=windows-1252", output, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("xcharset=utf-8", output, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SvgImageSrcIsNotAnActiveImageCarrier() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<html><body><svg xmlns=\"http://www.w3.org/2000/svg\"><image src=\"" +
            dataUri + "\"/></svg></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void XlsbWorkbookRelationshipResolvesDotSegments() {
        byte[] package = ReplaceWave38Entry(
            CreateWave33XlsbProvenancePackage(signed: false),
            "_rels/.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/worksheets/../workbook.bin\"/>" +
            "<Relationship Id=\"rApp\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
            "</Relationships>");

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
        try {
            File.WriteAllBytes(path, package);
            OfficeProvenanceReport report = ExcelDocument.InspectProvenance(path);

            Assert.Single(report.Evidence);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void XlsbApplicationOnlySignatureMetadataIsRemovedExplicitly() {
        byte[] package = ReplaceWave38Entry(
            CreateWave33XlsbProvenancePackage(signed: false),
            "docProps/app.xml",
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" +
            "<DigSig>present</DigSig></Properties>");
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeIMO.OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        OfficeProvenanceRemovalResult result = ExcelDocument.RemoveProvenance(package, "workbook.xlsb", options);
        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);

        Assert.True(result.WereInvalidatedSignaturesRemoved);
        Assert.DoesNotContain("DigSig", ReadWave33Entry(archive, "docProps/app.xml"), StringComparison.Ordinal);
    }
}
