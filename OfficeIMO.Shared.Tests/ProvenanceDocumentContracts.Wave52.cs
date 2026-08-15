using System.IO.Compression;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void MarkdownFileRemovalRejectsInvalidUtf8WithoutWritingOutput() {
        byte[] carrier = Encoding.ASCII.GetBytes(
            "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n");
        byte[] input = carrier.Concat(new byte[] { 0xFF }).ToArray();
        string inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
        try {
            File.WriteAllBytes(inputPath, input);

            Assert.Throws<InvalidDataException>(() => MarkdownProvenance.InspectFile(inputPath));
            Assert.Throws<InvalidDataException>(() => MarkdownProvenance.RemoveFile(inputPath, outputPath));
            Assert.False(File.Exists(outputPath));
            Assert.Equal(input, File.ReadAllBytes(inputPath));
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void HtmlStyleTypeUsesOnlyAsciiWhitespace() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<html><head><style type=\"\u00A0text/css\">.hero{background:url(" + dataUri + ")}</style></head>" +
            "<body><div class=\"hero\"></div></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void OdfSignatureManifestRewriteCountsOutputBytesOnlyOnce() {
        const string signaturePath = "META-INF/documentsignatures.xml";
        const string manifestPath = "META-INF/manifest.xml";
        byte[] cleanedManifest = Encoding.UTF8.GetBytes(
            "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\"/>");
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, "content.xml", "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>", CompressionLevel.Optimal);
                WriteEntry(archive, signaturePath, "<signatures/>", CompressionLevel.Optimal);
                WriteEntry(archive, manifestPath,
                    "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">" +
                    "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/>" +
                    "<manifest:file-entry manifest:full-path=\"" + signaturePath + "\" manifest:media-type=\"text/xml\"/>" +
                    "</manifest:manifest>", CompressionLevel.Optimal);
            }
            package = RewriteFixtureWithStoredMimetype(output.ToArray());
        }
        long expectedExpandedBytes = cleanedManifest.LongLength;
        using (var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read)) {
            expectedExpandedBytes += archive.Entries
                .Where(entry => entry.FullName is not signaturePath and not manifestPath)
                .Sum(entry => entry.Length);
        }

        OfficeProvenanceSignatureStripResult result = OfficeProvenanceZip.RemoveEntries(
            package,
            name => name == signaturePath,
            expectedExpandedBytes,
            name => name == manifestPath,
            (_, _) => cleanedManifest,
            maximumReplacementBytes: 4096);

        Assert.True(result.HadSignatures);
        Assert.Equal(cleanedManifest, ReadZipEntry(result.Data, manifestPath));
    }

    [Fact]
    public void ExcelXlsbRejectsMultipleInternalWorkbookRelationships() {
        byte[] package = ReplaceWave38Entry(
            CreateWave33XlsbProvenancePackage(signed: false),
            "_rels/.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/other.bin\"/>" +
            "</Relationships>");

        Assert.ThrowsAny<Exception>(() => ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));
    }
}
