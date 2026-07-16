using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
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
                    "<Override PartName=\"/xl/workbook.bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>" +
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
