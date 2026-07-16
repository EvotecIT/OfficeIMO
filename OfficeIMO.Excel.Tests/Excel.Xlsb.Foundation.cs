using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Excel.Xlsb;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Xlsb_RecordReader_FramesCanonicalWorkbookBoundaryRecords() {
            byte[] bytes = {
                0x83, 0x01, 0x00, // BrtBeginBook (131), empty payload
                0x84, 0x01, 0x00  // BrtEndBook (132), empty payload
            };

            using var stream = new MemoryStream(bytes, writable: false);
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(stream);

            Assert.Collection(records,
                begin => {
                    Assert.Equal(0, begin.Offset);
                    Assert.Equal(3, begin.HeaderSize);
                    Assert.Equal(131, begin.Type);
                    Assert.Empty(begin.Data);
                },
                end => {
                    Assert.Equal(3, end.Offset);
                    Assert.Equal(3, end.HeaderSize);
                    Assert.Equal(132, end.Type);
                    Assert.Empty(end.Data);
                });
        }

        [Fact]
        public void Xlsb_RecordReader_DecodesMultiByteTypeAndSize() {
            byte[] bytes = new byte[204];
            bytes[0] = 0xFD;
            bytes[1] = 0x04; // BrtCommentText (637)
            bytes[2] = 0xC8;
            bytes[3] = 0x01; // 200 payload bytes
            for (int index = 4; index < bytes.Length; index++) {
                bytes[index] = (byte)(index - 4);
            }

            using var stream = new MemoryStream(bytes, writable: false);
            XlsbRecord record = Assert.Single(XlsbRecordReader.ReadAll(stream));

            Assert.Equal(637, record.Type);
            Assert.Equal(200, record.Size);
            Assert.Equal(4, record.HeaderSize);
            Assert.Equal(0, record.Data[0]);
            Assert.Equal(199, record.Data[199]);
        }

        [Theory]
        [InlineData(new byte[] { 0x83 })]
        [InlineData(new byte[] { 0x83, 0x01, 0x80 })]
        [InlineData(new byte[] { 0x01, 0x02, 0xAA })]
        public void Xlsb_RecordReader_RejectsTruncatedRecords(byte[] bytes) {
            using var stream = new MemoryStream(bytes, writable: false);

            Assert.Throws<EndOfStreamException>(() => XlsbRecordReader.ReadAll(stream));
        }

        [Fact]
        public void Xlsb_RecordReader_EnforcesAllocationLimitBeforeReadingPayload() {
            using var stream = new MemoryStream(new byte[] { 0x01, 0x02, 0xAA, 0xBB }, writable: false);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                XlsbRecordReader.ReadAll(stream, maxRecordBytes: 1));

            Assert.Contains("configured limit of 1 byte", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Xlsb_PackageDetector_UsesPackageContentInsteadOfExtension() {
            byte[] package = CreateMinimalXlsbPackage();

            Assert.True(XlsbPackageDetector.TryFindWorkbookPart(package, out string? workbookPart));
            Assert.Equal("xl/workbook.bin", workbookPart);
            Assert.Equal(ExcelFileFormat.Xlsb, ExcelDocumentLoadRouting.DetectFormat(package, "misleading.xlsx"));
        }

        [Fact]
        public void Xlsb_PackageDetector_DoesNotMisclassifyXmlWorkbookPackages() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data").CellValue(1, 1, "XML package");
            byte[] package = document.ToBytes();

            Assert.False(XlsbPackageDetector.TryFindWorkbookPart(package, out _));
            Assert.Equal(ExcelFileFormat.Xlsx, ExcelDocumentLoadRouting.DetectFormat(package, "misleading.xlsb"));
        }

        [Fact]
        public void Xlsb_ExcelGeneratedFixture_HasCanonicalPackageAndWorkbookRecords() {
            string path = Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "XlsbCorpus",
                "excel-generated",
                "basic-values-formula.xlsb");
            byte[] package = File.ReadAllBytes(path);

            Assert.True(XlsbPackageDetector.TryFindWorkbookPart(package, out string? workbookPart));
            Assert.Equal("xl/workbook.bin", workbookPart);
            Assert.Equal(ExcelFileFormat.Xlsb, ExcelDocumentLoadRouting.DetectFormat(package, path));

            using var packageStream = new MemoryStream(package, writable: false);
            using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: false);
            ZipArchiveEntry workbookEntry = Assert.Single(
                archive.Entries,
                entry => string.Equals(entry.FullName, workbookPart, StringComparison.OrdinalIgnoreCase));
            using Stream workbookStream = workbookEntry.Open();
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(workbookStream);

            Assert.NotEmpty(records);
            Assert.Equal(131, records[0].Type); // BrtBeginBook
            Assert.Equal(132, records[records.Count - 1].Type); // BrtEndBook
        }

        [Fact]
        public void Xlsb_ExcelGeneratedFixture_LoadsThroughNormalExcelDocumentSurface() {
            string path = Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "XlsbCorpus",
                "excel-generated",
                "basic-values-formula.xlsb");

            using ExcelDocument document = ExcelDocument.Load(path);

            Assert.Equal(ExcelFileFormat.Xlsb, document.SourceFormat);
            Assert.Equal(path, document.SourcePath);
            ExcelSheet sheet = Assert.Single(document.Sheets);
            Assert.Equal("Arkusz1", sheet.Name);
            Assert.True(sheet.TryGetCellText(1, 1, out string? a1));
            Assert.True(sheet.TryGetCellText(1, 2, out string? b1));
            Assert.True(sheet.TryGetCellText(2, 1, out string? a2));
            Assert.True(sheet.TryGetCellText(2, 2, out string? b2));
            Assert.True(sheet.TryGetCellText(3, 2, out string? b3));
            Assert.Equal("Name", a1);
            Assert.Equal("Amount", b1);
            Assert.Equal("Alpha", a2);
            Assert.Equal("42", b2);
            Assert.Equal("50", b3);

            Cell formulaCell = Assert.Single(
                sheet.DeferredMetadataWorksheetPart.Worksheet.Descendants<Cell>(),
                cell => cell.CellReference?.Value == "B3");
            Assert.Equal("SUM(B2,8)", formulaCell.CellFormula?.Text);
            Assert.NotEmpty(document.XlsbPreservedRecords);
            Assert.Contains(document.XlsbImportDiagnostics, diagnostic => diagnostic.Code == "XLSB-RECORDS-PRESERVED");
        }

        [Fact]
        public void Xlsb_UnmodifiedSource_CopiesByteForByteToNativeTarget() {
            byte[] source = File.ReadAllBytes(GetExcelGeneratedXlsbFixturePath());
            using var input = new MemoryStream(source, writable: false);
            using ExcelDocument document = ExcelDocument.Load(input);

            byte[] saved = document.ToBytes(ExcelFileFormat.Xlsb);

            Assert.Equal(source, saved);
            Assert.Equal(ExcelFileFormat.Xlsb, ExcelDocumentLoadRouting.DetectFormat(saved, "copy.xlsb"));
        }

        [Fact]
        public void Xlsb_ModifiedSource_RejectsNativeRewriteBeforeWriting() {
            using ExcelDocument document = ExcelDocument.Load(GetExcelGeneratedXlsbFixturePath());
            document.Sheets[0].CellValue(2, 2, 43);
            using var destination = new MemoryStream();

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                document.Save(destination, ExcelFileFormat.Xlsb));

            Assert.Contains("modified", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, destination.Length);
        }

        [Fact]
        public void Xlsb_ProjectedWorkbook_SavesAsValidEditableXlsx() {
            using ExcelDocument source = ExcelDocument.Load(GetExcelGeneratedXlsbFixturePath());
            source.Sheets[0].CellValue(2, 2, 43);
            using var destination = new MemoryStream();

            source.Save(destination, ExcelFileFormat.Xlsx);
            byte[] xlsx = destination.ToArray();

            Assert.Equal(ExcelFileFormat.Xlsx, ExcelDocumentLoadRouting.DetectFormat(xlsx, "converted.xlsx"));
            using ExcelDocument converted = ExcelDocument.Load(new MemoryStream(xlsx, writable: false));
            Assert.True(converted.Sheets[0].TryGetCellText(2, 2, out string? value));
            Assert.Equal("43", value);
            Cell formulaCell = Assert.Single(
                converted.Sheets[0].DeferredMetadataWorksheetPart.Worksheet.Descendants<Cell>(),
                cell => cell.CellReference?.Value == "B3");
            Assert.Equal("SUM(B2,8)", formulaCell.CellFormula?.Text);
        }

        [Fact]
        public void Xlsb_ImportLimits_BlockCellExpansion() {
            var options = new ExcelLoadOptions {
                XlsbImportOptions = new XlsbImportOptions { MaxCells = 4 }
            };

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.Load(GetExcelGeneratedXlsbFixturePath(), options));

            Assert.Contains("limit of 4 populated cells", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task Xlsb_SaveTargets_NeverWriteMislabeledXlsxContent() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data").CellValue(1, 1, "XLSB");
            using var synchronousDestination = new MemoryStream();
            using var asynchronousDestination = new MemoryStream();

            Exception? synchronousFailure = Record.Exception(() =>
                document.Save(synchronousDestination, ExcelFileFormat.Xlsb));
            Exception? asynchronousFailure = await Record.ExceptionAsync(() =>
                document.SaveAsync(asynchronousDestination, ExcelFileFormat.Xlsb));

            AssertXlsbSaveResult(synchronousDestination, synchronousFailure);
            AssertXlsbSaveResult(asynchronousDestination, asynchronousFailure);
        }

        private static byte[] CreateMinimalXlsbPackage() {
            byte[] workbookRecords = {
                0x83, 0x01, 0x00,
                0x84, 0x01, 0x00
            };

            using var package = new MemoryStream();
            using (var archive = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteZipEntry(
                    archive,
                    "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>" +
                    "</Types>");
                WriteZipEntry(
                    archive,
                    "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
                    "</Relationships>");
                WriteZipEntry(archive, "xl/workbook.bin", workbookRecords);
            }

            return package.ToArray();
        }

        private static string GetExcelGeneratedXlsbFixturePath() {
            return Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "XlsbCorpus",
                "excel-generated",
                "basic-values-formula.xlsb");
        }

        private static void WriteZipEntry(ZipArchive archive, string name, string content) {
            WriteZipEntry(archive, name, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(content));
        }

        private static void WriteZipEntry(ZipArchive archive, string name, byte[] content) {
            ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Fastest);
            using Stream stream = entry.Open();
            stream.Write(content, 0, content.Length);
        }

        private static void AssertXlsbSaveResult(MemoryStream destination, Exception? failure) {
            if (failure != null) {
                Assert.IsType<NotSupportedException>(failure);
                Assert.Equal(0, destination.Length);
                return;
            }

            Assert.Equal(
                ExcelFileFormat.Xlsb,
                ExcelDocumentLoadRouting.DetectFormat(destination.ToArray(), "workbook.xlsb"));
        }
    }
}
