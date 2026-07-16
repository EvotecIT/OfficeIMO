using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Word;
using System.IO.Compression;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingOfficePackageSecurityTests {
    [Fact]
    public void InspectorInventoriesActiveEmbeddedExternalAndSignedContent() {
        byte[] package = CreateZip(archive => {
            AddEntry(archive, "[Content_Types].xml", "<Types />");
            AddEntry(archive, "word/vbaProject.bin", "vba");
            AddEntry(archive, "word/embeddings/oleObject1.bin", "ole");
            AddEntry(archive, "word/activeX/activeX1.bin", "control");
            AddEntry(archive, "_xmlsignatures/sig1.xml", "<Signature />");
            AddEntry(archive, "word/_rels/document.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"urn:test\" Target=\"https://example.test/item\" TargetMode=\"External\" />" +
                "</Relationships>");
        });

        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(package);

        Assert.True(report.IsValid);
        Assert.Equal(OfficePackageContainerKind.OpenXml, report.ContainerKind);
        Assert.Equal(1, report.MacroPartCount);
        Assert.Equal(1, report.EmbeddedPayloadPartCount);
        Assert.Equal(1, report.ActiveXPartCount);
        Assert.Equal(1, report.ExternalRelationshipCount);
        Assert.Equal(1, report.DigitalSignaturePartCount);
    }

    [Fact]
    public void UntrustedPolicyRejectsEachActiveContentClass() {
        byte[] package = CreateZip(archive => {
            AddEntry(archive, "word/vbaProject.bin", "vba");
            AddEntry(archive, "word/embeddings/package1.bin", "package");
            AddEntry(archive, "word/activeX/activeX1.bin", "control");
            AddEntry(archive, "_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"urn:test\" Target=\"file:///tmp/item\" TargetMode=\"External\" />" +
                "</Relationships>");
        });

        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(package,
            OfficePackageSecurityOptions.UntrustedDefaults);

        Assert.False(report.IsValid);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.Macros);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.EmbeddedPayloads);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.ActiveX);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.ExternalRelationships);
    }

    [Fact]
    public void UntrustedPolicyClassifiesActiveContentByContentTypeAndRelationshipType() {
        byte[] package = CreateZip(archive => {
            AddEntry(archive, "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Override PartName=\"/custom/macro.bin\" ContentType=\"application/vnd.ms-office.vbaProject\" />" +
                "</Types>");
            AddEntry(archive, "custom/macro.bin", "vba");
            AddEntry(archive, "payload/ole.bin", "ole");
            AddEntry(archive, "payload/control.bin", "control");
            AddEntry(archive, "_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rOle\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject\" Target=\"payload/ole.bin\" />" +
                "<Relationship Id=\"rControl\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/control\" Target=\"payload/control.bin\" />" +
                "</Relationships>");
        });

        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(
            package,
            OfficePackageSecurityOptions.UntrustedDefaults);

        Assert.False(report.IsValid);
        Assert.Equal(1, report.MacroPartCount);
        Assert.Equal(1, report.EmbeddedPayloadPartCount);
        Assert.Equal(1, report.ActiveXPartCount);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.Macros);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.EmbeddedPayloads);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.ActiveX);
    }

    [Fact]
    public void ValidatorRejectsHighlyCompressedPartsBeforeOpeningThem() {
        byte[] package = CreateZip(archive => {
            ZipArchiveEntry entry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Optimal);
            using Stream output = entry.Open();
            output.Write(new byte[256 * 1024], 0, 256 * 1024);
        });
        var options = OfficePackageSecurityOptions.SecureDefaults;
        options.MaxCompressionRatio = 10;

        OfficePackageSecurityException exception = Assert.Throws<OfficePackageSecurityException>(() =>
            OfficePackageSecurityInspector.Validate(package, options));

        Assert.Equal(OfficePackageSecurityRule.CompressionRatio, exception.Rule);
        Assert.Equal("/xl/worksheets/sheet1.xml", exception.PartName);
        Assert.True(exception.ObservedValue > exception.Limit);
    }

    [Fact]
    public void InspectorDoesNotParseRelationshipPartsThatExceedCompressionLimits() {
        byte[] package = CreateZip(archive =>
            AddEntry(archive, "_rels/.rels", new string('A', 256 * 1024)));
        var options = OfficePackageSecurityOptions.SecureDefaults;
        options.MaxCompressionRatio = 10;

        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(package, options);

        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.CompressionRatio);
        Assert.DoesNotContain(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.MalformedRelationship);
        Assert.Equal(0, report.ExternalRelationshipCount);
    }

    [Fact]
    public void InspectorCountsDirectoryEntriesBeforeMaterializingZipParts() {
        byte[] package = CreateZip(archive => {
            archive.CreateEntry("one/");
            archive.CreateEntry("two/");
            archive.CreateEntry("three/");
        });
        var options = OfficePackageSecurityOptions.SecureDefaults;
        options.MaxPartCount = 2;

        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(package, options);

        OfficePackageSecurityFinding finding = Assert.Single(
            report.Findings,
            item => item.Rule == OfficePackageSecurityRule.PartCount);
        Assert.Equal(3D, finding.ObservedValue);
        Assert.Equal(2D, finding.Limit);
        Assert.Equal(0, report.PartCount);
    }

    [Fact]
    public void InspectorReportsAmbiguousUnsafeAndMalformedRelationshipParts() {
        byte[] package = CreateZip(archive => {
            AddEntry(archive, "word/document.xml", "<document />");
            AddEntry(archive, "WORD/document.xml", "<document />");
            AddEntry(archive, "../escape.bin", "escape");
            AddEntry(archive, "word/_rels/document.xml.rels", "<Relationships><Relationship");
        });

        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(package);

        Assert.False(report.IsValid);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.DuplicatePartName);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.UnsafePartName);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.MalformedRelationship);
    }

    [Fact]
    public void StreamInspectionPreservesPositionAndThrowsTypedSourceLimit() {
        using var source = new MemoryStream(new byte[64]);
        source.Position = 17;
        var options = OfficePackageSecurityOptions.SecureDefaults;
        options.MaxPackageBytes = 32;

        OfficePackageSecurityException exception = Assert.Throws<OfficePackageSecurityException>(() =>
            OfficePackageSecurityInspector.Inspect(source, options));

        Assert.Equal(OfficePackageSecurityRule.PackageSize, exception.Rule);
        Assert.Equal(17, source.Position);
    }

    [Fact]
    public void InspectorAppliesActiveContentPoliciesToLegacyCompoundFiles() {
        var streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
            ["Workbook"] = new byte[] { 1, 2, 3 },
            ["Macros/VBA/dir"] = new byte[] { 4, 5 },
            ["ObjectPool/Item/Contents"] = new byte[] { 6, 7 },
            ["ObjectPool/Item/OCXNAME"] = new byte[] { 8 }
        };
        var entries = streams.Select(stream =>
            new OfficeCompoundFileEntry(Path.GetFileName(stream.Key), stream.Key, 2, stream.Value.Length))
            .ToArray();
        var compound = new OfficeCompoundFile(streams, entries,
            new OfficeCompoundFileEntry("Root Entry", "Root Entry", 5, 0));
        byte[] package = OfficeCompoundFileWriter.Rewrite(compound,
            new Dictionary<string, byte[]>());

        OfficePackageSecurityReport report = OfficePackageSecurityInspector.Inspect(package,
            OfficePackageSecurityOptions.UntrustedDefaults);

        Assert.Equal(OfficePackageContainerKind.CompoundBinary, report.ContainerKind);
        Assert.Equal(8, report.PartCount);
        Assert.Equal(1, report.MacroPartCount);
        Assert.Equal(2, report.EmbeddedPayloadPartCount);
        Assert.Equal(1, report.ActiveXPartCount);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.Macros);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.EmbeddedPayloads);
        Assert.Contains(report.Findings, finding => finding.Rule == OfficePackageSecurityRule.ActiveX);
    }

    [Fact]
    public void WordAndExcelRejectLegacyActiveContentBeforeFormatParsing() {
        var streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
            ["Workbook"] = new byte[] { 1, 2, 3 },
            ["Macros/VBA/dir"] = new byte[] { 4, 5 }
        };
        var entries = streams.Select(stream =>
            new OfficeCompoundFileEntry(Path.GetFileName(stream.Key), stream.Key, 2, stream.Value.Length))
            .ToArray();
        var compound = new OfficeCompoundFile(streams, entries,
            new OfficeCompoundFileEntry("Root Entry", "Root Entry", 5, 0));
        byte[] package = OfficeCompoundFileWriter.Rewrite(compound,
            new Dictionary<string, byte[]>());
        OfficePackageSecurityOptions security = OfficePackageSecurityOptions.UntrustedDefaults;

        using var excelSource = new MemoryStream(package);
        OfficePackageSecurityException excelException = Assert.Throws<OfficePackageSecurityException>(() =>
            ExcelDocument.Load(excelSource, new ExcelLoadOptions { PackageSecurity = security }));
        Assert.Equal(OfficePackageSecurityRule.Macros, excelException.Rule);

        using var wordSource = new MemoryStream(package);
        OfficePackageSecurityException wordException = Assert.Throws<OfficePackageSecurityException>(() =>
            WordDocument.Load(wordSource, new WordLoadOptions { PackageSecurity = security }));
        Assert.Equal(OfficePackageSecurityRule.Macros, wordException.Rule);
    }

    [Fact]
    public void WordAndExcelLoadsUseTheSamePartCountPolicy() {
        byte[] workbook = CreateWorkbook();
        byte[] document = CreateDocument();
        var options = OfficePackageSecurityOptions.SecureDefaults;
        options.MaxPartCount = 1;

        using var workbookStream = new MemoryStream(workbook);
        workbookStream.Position = 3;
        OfficePackageSecurityException excelException = Assert.Throws<OfficePackageSecurityException>(() =>
            ExcelDocument.Load(workbookStream, new ExcelLoadOptions { PackageSecurity = options }));
        Assert.Equal(OfficePackageSecurityRule.PartCount, excelException.Rule);
        Assert.Equal(3, workbookStream.Position);

        using var documentStream = new MemoryStream(document);
        documentStream.Position = 5;
        OfficePackageSecurityException wordException = Assert.Throws<OfficePackageSecurityException>(() =>
            WordDocument.Load(documentStream, new WordLoadOptions { PackageSecurity = options }));
        Assert.Equal(OfficePackageSecurityRule.PartCount, wordException.Rule);
        Assert.Equal(5, documentStream.Position);
    }

    [Fact]
    public void WordAndExcelLoadsRejectExternalRelationshipsWhenRequested() {
        byte[] workbook = CreateWorkbook();
        using (var editable = CreateExpandableStream(workbook)) {
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(editable, true)) {
                package.WorkbookPart!.AddExternalRelationship("urn:officeimo:test",
                    new Uri("https://example.test/workbook"), "rExternal");
            }
            workbook = editable.ToArray();
        }

        byte[] document = CreateDocument();
        using (var editable = CreateExpandableStream(document)) {
            using (WordprocessingDocument package = WordprocessingDocument.Open(editable, true)) {
                package.MainDocumentPart!.AddExternalRelationship("urn:officeimo:test",
                    new Uri("https://example.test/document"), "rExternal");
            }
            document = editable.ToArray();
        }

        OfficePackageSecurityOptions security = OfficePackageSecurityOptions.UntrustedDefaults;
        using var workbookStream = new MemoryStream(workbook);
        OfficePackageSecurityException excelException = Assert.Throws<OfficePackageSecurityException>(() =>
            ExcelDocument.Load(workbookStream, new ExcelLoadOptions { PackageSecurity = security }));
        Assert.Equal(OfficePackageSecurityRule.ExternalRelationships, excelException.Rule);

        using var documentStream = new MemoryStream(document);
        OfficePackageSecurityException wordException = Assert.Throws<OfficePackageSecurityException>(() =>
            WordDocument.Load(documentStream, new WordLoadOptions { PackageSecurity = security }));
        Assert.Equal(OfficePackageSecurityRule.ExternalRelationships, wordException.Rule);
    }

    [Fact]
    public async Task WordAndExcelAsyncLoadsUseTheSameSourceSizePolicy() {
        byte[] workbook = CreateWorkbook();
        byte[] document = CreateDocument();
        var options = OfficePackageSecurityOptions.SecureDefaults;
        options.MaxPackageBytes = 128;

        using var workbookStream = new MemoryStream(workbook);
        OfficePackageSecurityException excelException = await Assert.ThrowsAsync<OfficePackageSecurityException>(() =>
            ExcelDocument.LoadAsync(workbookStream, new ExcelLoadOptions { PackageSecurity = options }));
        Assert.Equal(OfficePackageSecurityRule.PackageSize, excelException.Rule);

        using var documentStream = new MemoryStream(document);
        OfficePackageSecurityException wordException = await Assert.ThrowsAsync<OfficePackageSecurityException>(() =>
            WordDocument.LoadAsync(documentStream, new WordLoadOptions { PackageSecurity = options }));
        Assert.Equal(OfficePackageSecurityRule.PackageSize, wordException.Rule);
    }

    private static byte[] CreateWorkbook() {
        using var output = new MemoryStream();
        using (ExcelDocument workbook = ExcelDocument.Create()) {
            workbook.AddWorksheet("Data").CellValue(1, 1, "safe");
            workbook.Save(output);
        }
        return output.ToArray();
    }

    private static byte[] CreateDocument() {
        using var output = new MemoryStream();
        using (WordDocument document = WordDocument.Create()) {
            document.AddParagraph("safe");
            document.Save(output);
        }
        return output.ToArray();
    }

    private static byte[] CreateZip(Action<ZipArchive> populate) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            populate(archive);
        }
        return output.ToArray();
    }

    private static MemoryStream CreateExpandableStream(byte[] bytes) {
        var stream = new MemoryStream(bytes.Length + 4096);
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        return stream;
    }

    private static void AddEntry(ZipArchive archive, string name, string content) {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
        writer.Write(content);
    }
}
