using System.Data.Common;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Read;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private const string XlsLatinText = "Café Ø ÿ";
        private const string XlsUnicodeText = "Zażółć \U0001F600";

        [Fact]
        public void OpenDataReader_XlsFastPathDecodesInlineStringEncodingsAndSkipsStringMetadata() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateInlineStringEncodingWorkbookStream());
            var options = new ExcelReaderOptions { HasHeaderRow = false };

            using (LegacyXlsTabularWorkbook workbook = LegacyXlsTabularWorkbook.Open(compound, options))
            using (var direct = Assert.IsType<LegacyXlsTabularDataReader>(
                       workbook.OpenTable("Strings", hasHeaderRow: false, options))) {
                AssertInlineStringEncodingRow(direct);
            }

            using DbDataReader reader = ExcelDocument.OpenReader(compound, options);
            AssertInlineStringEncodingRow(reader);
        }

        private static void AssertInlineStringEncodingRow(DbDataReader reader) {
            Assert.Equal(8, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.Equal(XlsLatinText, reader.GetString(0));
            Assert.Equal(XlsUnicodeText, reader.GetString(1));
            Assert.Equal("Rich", reader.GetString(2));
            Assert.Equal("Mixed", reader.GetString(3));
            Assert.Equal(XlsUnicodeText, reader.GetString(4));
            Assert.Equal("Latin then Ωmega", reader.GetString(5));
            Assert.Equal("Shared Ωmega", reader.GetString(6));
            Assert.Equal("Next", reader.GetString(7));
            Assert.False(reader.Read());
        }

        [Fact]
        public void OpenDataReader_XlsFastPathCopiesCharacterRanges() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateInlineStringEncodingWorkbookStream());
            var options = new ExcelReaderOptions { HasHeaderRow = false };

            using LegacyXlsTabularWorkbook workbook = LegacyXlsTabularWorkbook.Open(compound, options);
            using var reader = Assert.IsType<LegacyXlsTabularDataReader>(
                workbook.OpenTable("Strings", hasHeaderRow: false, options));
            Assert.True(reader.Read());

            for (int ordinal = 0; ordinal < 2; ordinal++) {
                string expected = reader.GetString(ordinal);
                Assert.Equal(expected.Length, reader.GetChars(ordinal, 0, null, 0, 0));

                var copied = new char[expected.Length + 2];
                long offset = 0;
                long read;
                while ((read = reader.GetChars(ordinal, offset, copied, checked((int)offset + 2), 3)) > 0) {
                    offset += read;
                }

                Assert.Equal(expected.Length, offset);
                Assert.Equal(expected, new string(copied, 2, expected.Length));
                Assert.Equal(0, reader.GetChars(ordinal, expected.Length, copied, 0, 3));
                Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetChars(ordinal, expected.Length + 1, copied, 0, 1));
                Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetChars(ordinal, -1, copied, 0, 1));
            }
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            private const byte HighByteOption = 0x01;
            private const byte ExtendedOption = 0x04;
            private const byte RichTextOption = 0x08;

            /// <summary>
            /// One row of inline LABEL, cached formula STRING and SST cells covering both BIFF
            /// character encodings, string metadata, the BIFF5-style byte-string label and
            /// strings whose characters continue into a CONTINUE record in another encoding.
            /// </summary>
            internal static byte[] CreateInlineStringEncodingWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Strings"));

                // "Shared " ends the SST record; the CONTINUE switches to UTF-16 for "Ωmega"
                // and then holds the next whole string.
                using (var sst = new MemoryStream()) {
                    WriteUInt32(sst, 2);
                    WriteUInt32(sst, 2);
                    WriteUInt16(sst, (ushort)"Shared Ωmega".Length);
                    sst.WriteByte(0);
                    WriteBytes(sst, Latin1("Shared "));
                    WriteRecord(stream, 0x00fc, sst.ToArray());
                }
                using (var continuation = new MemoryStream()) {
                    continuation.WriteByte(HighByteOption);
                    WriteBytes(continuation, Encoding.Unicode.GetBytes("Ωmega"));
                    WriteBiffString(continuation, "Next", options: 0);
                    WriteRecord(stream, 0x003c, continuation.ToArray());
                }
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildInlineLabelPayload(0, w => WriteBiffString(w, XlsLatinText, 0)));
                WriteRecord(stream, 0x0204, BuildInlineLabelPayload(1, w => WriteBiffString(w, XlsUnicodeText, HighByteOption)));
                WriteRecord(stream, 0x0204, BuildInlineLabelPayload(2, w => {
                    WriteUInt16(w, 4);
                    w.WriteByte(RichTextOption | ExtendedOption);
                    WriteUInt16(w, 1);
                    WriteUInt32(w, 3);
                    WriteBytes(w, Latin1("Rich"));
                    WriteBytes(w, new byte[] { 0x00, 0x00, 0x01, 0x00 });
                    WriteBytes(w, new byte[] { 0xaa, 0xbb, 0xcc });
                }));
                // A two-byte length and bytes without an option byte: the leading 'M' reads
                // as options that overrun the record, so the label is read as a byte string.
                WriteRecord(stream, 0x0204, BuildInlineLabelPayload(3, w => {
                    WriteUInt16(w, 5);
                    WriteBytes(w, Latin1("Mixed"));
                }));
                WriteRecord(stream, 0x0006, BuildFormulaSpecialPayload(0, 4, valueType: 0x00, value: 0));
                using (var formulaString = new MemoryStream()) {
                    WriteBiffString(formulaString, XlsUnicodeText, HighByteOption);
                    WriteRecord(stream, 0x0207, formulaString.ToArray());
                }
                WriteRecord(stream, 0x0006, BuildFormulaSpecialPayload(0, 5, valueType: 0x00, value: 0));
                using (var formulaString = new MemoryStream()) {
                    WriteUInt16(formulaString, (ushort)"Latin then Ωmega".Length);
                    formulaString.WriteByte(0);
                    WriteBytes(formulaString, Latin1("Latin then "));
                    WriteRecord(stream, 0x0207, formulaString.ToArray());
                }
                using (var continuation = new MemoryStream()) {
                    continuation.WriteByte(HighByteOption);
                    WriteBytes(continuation, Encoding.Unicode.GetBytes("Ωmega"));
                    WriteRecord(stream, 0x003c, continuation.ToArray());
                }
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 6, 0));
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 7, 1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            private static byte[] BuildInlineLabelPayload(ushort column, Action<Stream> writeString) {
                using var payload = new MemoryStream();
                WriteUInt16(payload, 0);
                WriteUInt16(payload, column);
                WriteUInt16(payload, 0);
                writeString(payload);
                return payload.ToArray();
            }

            private static void WriteBiffString(Stream stream, string text, int options) {
                WriteUInt16(stream, checked((ushort)text.Length));
                stream.WriteByte((byte)options);
                byte[] characters = (options & HighByteOption) != 0 ? Encoding.Unicode.GetBytes(text) : Latin1(text);
                WriteBytes(stream, characters);
            }

            private static byte[] Latin1(string text) {
                var bytes = new byte[text.Length];
                for (int index = 0; index < text.Length; index++) {
                    bytes[index] = checked((byte)text[index]);
                }
                return bytes;
            }

            private static void WriteBytes(Stream stream, byte[] bytes) => stream.Write(bytes, 0, bytes.Length);
        }
    }
}
