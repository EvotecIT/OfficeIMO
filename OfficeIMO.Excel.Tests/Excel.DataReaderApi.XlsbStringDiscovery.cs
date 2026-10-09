using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests {

    public sealed class XlsbStringDiscoveryTests {
        [Theory]
        [InlineData(6, false)]
        [InlineData(62, false)]
        [InlineData(8, false)]
        [InlineData(6, true)]
        [InlineData(62, true)]
        [InlineData(8, true)]
        public void NativeXlsbStringsPreserveDecoderPolicyAndMaterializedLifetime(int recordType, bool empty) {
            // Include a valid surrogate pair, unmatched surrogates and a NUL without
            // passing them through an encoder that would replace the malformed units.
            ushort[] units = empty ? Array.Empty<ushort>()
                : new ushort[] { 0x41, 0xD800, 0x42, 0x20, 0xD83D, 0xDC22, 0, 0x5A, 0xDC00 };
            string expected = empty ? string.Empty : "A\uFFFDB 🐢\0Z\uFFFD";
            byte[] workbook = CreateStringDiscoveryWorkbook(recordType, units);
            string materialized;
            using (ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook)) {
                Assert.True(reader.Read());
                Assert.Equal("Ready", reader.GetString(0));
                Assert.True(reader.Read());
                Assert.False(reader.IsDBNull(0));
                materialized = reader.GetString(0);
                Assert.Equal(expected, materialized);
                Assert.Equal(expected, Assert.IsType<string>(reader.GetValue(0)));
                Assert.Equal(43, reader.GetInt32(1));
                Assert.True(reader.Read());
                Assert.Equal("After", reader.GetString(0));
                Assert.Equal(expected, materialized);
                Assert.False(reader.Read());
            }
            Assert.Equal(expected, materialized);
        }

        [Theory]
        [InlineData(6, "Count")]
        [InlineData(62, "Count")]
        [InlineData(8, "Count")]
        [InlineData(6, "Truncated")]
        [InlineData(62, "Truncated")]
        [InlineData(8, "Truncated")]
        [InlineData(8, "MissingTail")]
        [InlineData(8, "TokenCount")]
        public void NativeXlsbRejectsLateStringPayloadFailuresBeforeReaderDelivery(int recordType, string failure) {
            byte[] workbook = CreateStringDiscoveryWorkbook(recordType, new ushort[] { 0x41, 0x42, 0x43 }, failure);
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(workbook));
            string expected = failure switch {
                "Count" => "exceeding the configured limit",
                "Truncated" => "truncated",
                _ => "token"
            };
            Assert.Contains(expected, error.Message, StringComparison.OrdinalIgnoreCase);
        }

        private static byte[] CreateStringDiscoveryWorkbook(int recordType, ushort[] units, string? failure = null) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Text");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Ready");
            sheet.CellValue(2, 2, 42);
            sheet.CellValue(3, 1, "Replace");
            sheet.CellValue(3, 2, 43);
            sheet.CellValue(4, 1, "After");
            sheet.CellValue(4, 2, 44);
            byte[] original = document.ToBytes(ExcelFileFormat.Xlsb, new ExcelSaveOptions { XlsbUseSharedStrings = false });
            using MemoryStream package = new MemoryStream();
            package.Write(original, 0, original.Length);
            using (ZipArchive archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
                ZipArchiveEntry part = archive.GetEntry("xl/worksheets/sheet1.bin")!;
                IReadOnlyList<XlsbRecord> records;
                using (Stream input = part.Open()) records = XlsbRecordReader.ReadAll(input);
                part.Delete();
                using Stream output = archive.CreateEntry("xl/worksheets/sheet1.bin").Open();
                int row = -1;
                bool replaced = false;
                foreach (XlsbRecord record in records) {
                    if (record.Type == 0) row = BitConverter.ToInt32(record.Data, 0);
                    if (row == 2 && record.Type == 6 && BitConverter.ToInt32(record.Data, 0) == 0) {
                        XlsbRecordWriter.Write(output, recordType, CreateStringDiscoveryPayload(record.Data, recordType, units, failure));
                        replaced = true;
                    } else {
                        XlsbRecordWriter.Write(output, record.Type, record.Data);
                    }
                }
                Assert.True(replaced);
            }
            return package.ToArray();
        }

        private static byte[] CreateStringDiscoveryPayload(byte[] original, int recordType, ushort[] units, string? failure) {
            using MemoryStream payload = new MemoryStream();
            using (BinaryWriter writer = new BinaryWriter(payload, Encoding.Unicode, leaveOpen: true)) {
                writer.Write(original, 0, 8); // Column and style.
                if (recordType == 62) writer.Write((byte)0); // Rich-string flags.
                writer.Write(failure == "Count"
                    ? checked((uint)new XlsbImportOptions().MaxStringCharacters + 1)
                    : (uint)units.Length);
                foreach (ushort unit in units) writer.Write(unit);
                if (recordType == 8 && failure != "MissingTail") {
                    writer.Write((ushort)0); // Formula flags.
                    writer.Write(failure == "TokenCount" ? 1U : 0U);
                }
            }
            byte[] bytes = payload.ToArray();
            return failure == "Truncated"
                ? bytes.Take(8 + (recordType == 62 ? 1 : 0) + sizeof(uint) + units.Length * sizeof(ushort) - 1).ToArray()
                : bytes;
        }
    }
}
