using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void ExcelXlsbSignatureCleanupIgnoresForeignContentTypeOverrides() {
        byte[] package = CreateWave33XlsbProvenancePackage(signed: true);
        package = ReplaceWave38Entry(
            package,
            "[Content_Types].xml",
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>" +
            "<Override xmlns=\"urn:foreign:content-types\" PartName=\"/xl/workbook.bin\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-origin\"/>" +
            "<Override PartName=\"/_xmlsignatures/origin.sigs\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-origin\"/>" +
            "<Override PartName=\"/_xmlsignatures/sig1.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "</Types>");
        package = ReplaceWave38Entry(
            package,
            "_rels/.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
            "<Relationship Id=\"rSig\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"xl/workbook.bin\"/>" +
            "</Relationships>");
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.RemoveProvenance(package, "workbook.xlsb", options));

        Assert.Contains("signature-origin target", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlScriptTypeUsesAsciiWhitespaceTrimming() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = $"<html><head><script type=\"\u00A0application/c2pa\">{manifest}</script></head><body></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(html, Encoding.UTF8.GetString(result.ToArray()));
    }

    [Fact]
    public void VisioSignatureCleanupResolvesNoncanonicalExtendedProperties() {
        byte[] package = CreateVisioPackageWithNoncanonicalApplicationSignatureMetadata();
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        OfficeProvenanceRemovalResult result = VisioDocument.RemoveProvenance(package, "drawing.vsdx", options);

        Assert.True(result.WasChanged);
        Assert.True(result.WereInvalidatedSignaturesRemoved);
        Assert.DoesNotContain("DigSig", Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "metadata/application.xml")), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MarkdownFileProvenancePreservesUtf16Encoding(bool bigEndian) {
        Encoding encoding = bigEndian ? Encoding.BigEndianUnicode : Encoding.Unicode;
        string block = "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n";
        string markdown = "before\n" + block + "after\n";
        byte[] preamble = encoding.GetPreamble();
        byte[] body = encoding.GetBytes(markdown);
        byte[] input = preamble.Concat(body).ToArray();
        string inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".md");
        try {
            File.WriteAllBytes(inputPath, input);

            OfficeProvenanceReport report = MarkdownProvenance.InspectFile(inputPath);
            OfficeProvenanceRemovalResult result = MarkdownProvenance.RemoveFile(inputPath, outputPath);
            byte[] output = File.ReadAllBytes(outputPath);

            Assert.True(report.HasC2paManifest);
            Assert.True(result.WasChanged);
            Assert.Equal(preamble, output.Take(preamble.Length));
            Assert.Equal("before\nafter\n", encoding.GetString(output, preamble.Length, output.Length - preamble.Length));
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void OdfInspectionAllowsButMutationRejectsEncryptedPackages() {
        const string manifest =
            "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">" +
            "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/>" +
            "<manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\">" +
            "<manifest:encryption-data/></manifest:file-entry></manifest:manifest>";
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, "META-INF/manifest.xml", manifest, CompressionLevel.Optimal);
                WriteEntry(archive, "content.xml", "<encrypted/>", CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/content_credential.c2pa", CreateManifestStore(), CompressionLevel.Optimal);
            }
            package = RewriteFixtureWithStoredMimetype(output.ToArray());
        }
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".odt");
        try {
            File.WriteAllBytes(path, package);
            Assert.True(OdfDocument.InspectProvenance(path).HasC2paManifest);
            Assert.Throws<OdfEncryptedPackageException>(() => OdfDocument.RemoveProvenance(package, "document.odt"));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OdfOwnershipUsesThePhysicallyLeadingMimetypeEntry() {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, "META-INF/manifest.xml", ValidOdfManifestXml, CompressionLevel.Optimal);
                WriteEntry(archive, "content.xml", "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>", CompressionLevel.Optimal);
            }
            package = RotateFirstCentralDirectoryRecordToEnd(RewriteFixtureWithStoredMimetype(output.ToArray()));
        }

        OfficeProvenanceRemovalResult result = OdfDocument.RemoveProvenance(package, "document.odt");

        Assert.False(result.WasChanged);
        Assert.Equal(package, result.ToArray());
    }

    private static byte[] CreateVisioPackageWithNoncanonicalApplicationSignatureMetadata() {
        byte[] original = CreateSignedVisioProvenancePackage(0);
        using var output = new MemoryStream();
        output.Write(original, 0, original.Length);
        output.Position = 0;
        using (Package package = Package.Open(output, FileMode.Open, FileAccess.ReadWrite)) {
            foreach (PackageRelationship relationship in package.GetRelationshipsByType(
                "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin").ToArray()) {
                package.DeleteRelationship(relationship.Id);
            }
            foreach (string name in new[] { "/_xmlsignatures/sig1.xml", "/_xmlsignatures/origin.sigs", "/docProps/app.xml" }) {
                Uri uri = PackUriHelper.CreatePartUri(new Uri(name, UriKind.Relative));
                if (package.PartExists(uri)) package.DeletePart(uri);
            }
            Uri customUri = PackUriHelper.CreatePartUri(new Uri("/metadata/application.xml", UriKind.Relative));
            PackagePart application = package.CreatePart(
                customUri,
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                CompressionOption.Maximum);
            using (var writer = new StreamWriter(application.GetStream(), new UTF8Encoding(false), 4096, leaveOpen: false)) {
                writer.Write("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><DigSig>signature</DigSig></Properties>");
            }
            package.CreateRelationship(
                customUri,
                TargetMode.Internal,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
        }
        return output.ToArray();
    }

    private static byte[] RotateFirstCentralDirectoryRecordToEnd(byte[] package) {
        int eocd = package.Length - 22;
        while (eocd >= 0 && BitConverter.ToUInt32(package, eocd) != 0x06054B50U) eocd--;
        if (eocd < 0) throw new InvalidDataException("ZIP end record was not found.");
        int centralOffset = checked((int)BitConverter.ToUInt32(package, eocd + 16));
        int count = BitConverter.ToUInt16(package, eocd + 10);
        var records = new List<byte[]>(count);
        int cursor = centralOffset;
        for (int index = 0; index < count; index++) {
            if (BitConverter.ToUInt32(package, cursor) != 0x02014B50U) throw new InvalidDataException("ZIP central directory is malformed.");
            int length = 46 + BitConverter.ToUInt16(package, cursor + 28) +
                BitConverter.ToUInt16(package, cursor + 30) + BitConverter.ToUInt16(package, cursor + 32);
            byte[] record = new byte[length];
            Buffer.BlockCopy(package, cursor, record, 0, length);
            records.Add(record);
            cursor += length;
        }
        if (records.Count < 2) throw new InvalidDataException("The fixture needs at least two ZIP entries.");
        byte[] result = (byte[])package.Clone();
        cursor = centralOffset;
        foreach (byte[] record in records.Skip(1).Concat(records.Take(1))) {
            Buffer.BlockCopy(record, 0, result, cursor, record.Length);
            cursor += record.Length;
        }
        return result;
    }
}
