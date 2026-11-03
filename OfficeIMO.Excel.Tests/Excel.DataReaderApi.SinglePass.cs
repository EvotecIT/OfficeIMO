using System.IO.Compression;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        /// <summary>
        /// The readers validate each row as they reach it, so rows before a malformed one are
        /// returned and the failure surfaces from the Read call for the bad row.
        /// </summary>
        [Theory]
        [InlineData(ExcelFileFormat.Xlsb)]
        [InlineData(ExcelFileFormat.Xls)]
        public void OpenReader_ReturnsValidRowsBeforeAMalformedRowThenThrows(ExcelFileFormat format) {
            byte[] workbook;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Name");
                for (int row = 2; row <= 5; row++) {
                    sheet.Cell(row, 1).SetValue("Row " + row);
                }
                workbook = document.ToBytes(new ExcelSaveOptions { Format = format, XlsbUseSharedStrings = true });
            }
            // The fourth shared-string cell is the third data row.
            workbook = format == ExcelFileFormat.Xlsb
                ? ReplaceXlsbSharedStringIndex(workbook, occurrence: 4)
                : ReplaceXlsSharedStringIndex(workbook, occurrence: 4);

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(workbook);
            Assert.True(reader.Read());
            Assert.Equal("Row 2", reader.GetString(0));
            Assert.True(reader.Read());
            Assert.Equal("Row 3", reader.GetString(0));
            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => reader.Read());
            Assert.Contains("shared", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        private static byte[] ReplaceXlsbSharedStringIndex(byte[] workbook, int occurrence) {
            using var package = new MemoryStream();
            package.Write(workbook, 0, workbook.Length);
            using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
                ZipArchiveEntry entry = archive.GetEntry("xl/worksheets/sheet1.bin")!;
                IReadOnlyList<XlsbRecord> records;
                using (Stream input = entry.Open()) records = XlsbRecordReader.ReadAll(input);
                entry.Delete();
                using Stream output = archive.CreateEntry("xl/worksheets/sheet1.bin").Open();
                int seen = 0;
                foreach (XlsbRecord record in records) {
                    byte[] data = record.Data;
                    if (record.Type == 7 && ++seen == occurrence) {
                        data = (byte[])data.Clone();
                        BitConverter.GetBytes(uint.MaxValue).CopyTo(data, 8);
                    }
                    XlsbRecordWriter.Write(output, record.Type, data);
                }
                Assert.Equal(5, seen);
            }
            return package.ToArray();
        }

        private static byte[] ReplaceXlsSharedStringIndex(byte[] workbook, int occurrence) {
            byte[] stream = ReadCompoundStream(workbook, "Workbook");
            int seen = 0;
            for (int offset = 0; offset + 4 <= stream.Length;) {
                int type = stream[offset] | stream[offset + 1] << 8;
                int length = stream[offset + 2] | stream[offset + 3] << 8;
                if (type == 0x00fd && ++seen == occurrence) {
                    BitConverter.GetBytes(uint.MaxValue).CopyTo(stream, offset + 4 + 6);
                }
                offset += 4 + length;
            }
            Assert.Equal(5, seen);
            return LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(stream);
        }
    }
}
