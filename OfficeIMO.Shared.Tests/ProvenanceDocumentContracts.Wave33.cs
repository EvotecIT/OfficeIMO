using System.IO.Compression;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlIgnoresSourceImagesOutsidePictureElements() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><body><div><source src=\"{dataUri}\" srcset=\"{dataUri} 1x\"></div></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(html, Encoding.UTF8.GetString(result.ToArray()));
    }

    [Fact]
    public void HtmlPreflightKeepsMathMlGlyphCdataInForeignContent() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<html><head><script type=\"application/c2pa\">" + manifest +
            "</script></head><body><math><mi><mglyph><![CDATA[" +
            string.Concat(Enumerable.Repeat("<div></div>", 64)) +
            "]]></mglyph></mi></math></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html, new OfficeProvenanceOptions { MaxContainerEntries = 12 });

        Assert.Single(report.Evidence);
    }

    [Fact]
    public void HtmlPreflightModelsScriptDoubleEscapedState() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<html><head><script type=\"application/c2pa\">" + manifest +
            "</script></head><body><script><!--<script></script>" +
            string.Concat(Enumerable.Repeat("<div></div>", 64)) +
            "</script></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html, new OfficeProvenanceOptions { MaxContainerEntries = 8 });

        Assert.Single(report.Evidence);
    }

    [Fact]
    public void ExcelProvenanceSupportsXlsbPackagesAndSignatureCleanup() {
        byte[] package = CreateWave33XlsbProvenancePackage(signed: true);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
        try {
            File.WriteAllBytes(path, package);
            OfficeProvenanceReport report = ExcelDocument.InspectProvenance(path);
            var options = new OfficeProvenanceRemovalOptions {
                SignatureMutationPolicy = OfficeIMO.OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
            };

            OfficeProvenanceRemovalResult result = ExcelDocument.RemoveProvenance(package, "workbook.xlsb", options);
            using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);

            Assert.Single(report.Evidence);
            Assert.Empty(result.After.Evidence);
            Assert.True(result.WereInvalidatedSignaturesRemoved);
            Assert.DoesNotContain(archive.Entries, entry => entry.FullName.StartsWith("_xmlsignatures/", StringComparison.Ordinal));
            Assert.DoesNotContain("digital-signature/origin", ReadWave33Entry(archive, "_rels/.rels"), StringComparison.Ordinal);
            Assert.DoesNotContain("digital-signature", ReadWave33Entry(archive, "[Content_Types].xml"), StringComparison.Ordinal);
            Assert.DoesNotContain("DigSig", ReadWave33Entry(archive, "docProps/app.xml"), StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static byte[] CreateWave33XlsbProvenancePackage(bool signed) {
        string signatureTypes = signed
            ? "<Override PartName=\"/_xmlsignatures/origin.sigs\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-origin\"/>" +
              "<Override PartName=\"/_xmlsignatures/sig1.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>"
            : string.Empty;
        string signatureRelationship = signed
            ? "<Relationship Id=\"rSig\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"_xmlsignatures/origin.sigs\"/>"
            : string.Empty;
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteWave33Entry(archive, "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>" +
                "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
                signatureTypes + "</Types>");
            WriteWave33Entry(archive, "_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
                signatureRelationship + "</Relationships>");
            WriteWave33Entry(archive, "xl/workbook.bin", new byte[] { 0x83, 0x01, 0x00, 0x84, 0x01, 0x00 });
            WriteWave33Entry(archive, "META-INF/content_credential.c2pa", CreateManifestStore());
            WriteWave33Entry(archive, "docProps/app.xml",
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" +
                (signed ? "<DigSig>present</DigSig>" : string.Empty) + "</Properties>");
            if (signed) {
                WriteWave33Entry(archive, "_xmlsignatures/origin.sigs", Array.Empty<byte>());
                WriteWave33Entry(archive, "_xmlsignatures/_rels/origin.sigs.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature\" Target=\"sig1.xml\"/>" +
                    "</Relationships>");
                WriteWave33Entry(archive, "_xmlsignatures/sig1.xml",
                    "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo>" +
                    "<CanonicalizationMethod Algorithm=\"http://www.w3.org/2001/10/xml-exc-c14n#\"/>" +
                    "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\"/>" +
                    "</SignedInfo><SignatureValue>AA==</SignatureValue></Signature>");
            }
        }
        return output.ToArray();
    }

    private static void WriteWave33Entry(ZipArchive archive, string name, string content) =>
        WriteWave33Entry(archive, name, Encoding.UTF8.GetBytes(content));

    private static void WriteWave33Entry(ZipArchive archive, string name, byte[] content) {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using Stream stream = entry.Open();
        stream.Write(content, 0, content.Length);
    }

    private static string ReadWave33Entry(ZipArchive archive, string name) {
        using Stream stream = archive.GetEntry(name)!.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
