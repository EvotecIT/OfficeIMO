using System.Data.Common;
using System.Globalization;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void OpenDataReader_XlsFastPathFormatsMulRkDateHeaders() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateMulRkDateHeaderWorkbookStream());

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                compound,
                new ExcelReadOptions { Culture = CultureInfo.InvariantCulture });

            DateTime expected = DateTime.FromOADate(45293D);
            Assert.Equal(expected.ToString(CultureInfo.InvariantCulture), reader.GetName(0));
            Assert.True(reader.Read());
            Assert.Equal(1D, reader.GetDouble(0));
            Assert.False(reader.Read());
        }

        [Fact]
        public void OpenDataReader_XlsFastPathResolvesInheritedDateStyles() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateInheritedDateStyleWorkbookStream());

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                compound,
                new ExcelReadOptions { HasHeaderRow = false });

            Assert.True(reader.Read());
            Assert.Equal(DateTime.FromOADate(45293D), reader.GetDateTime(0));
            Assert.False(reader.Read());
        }

        [Fact]
        public void OpenDataReader_XlsFastPathKeepsInvalidDateSerialsNumeric() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateInvalidDateSerialWorkbookStream());

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                compound,
                new ExcelReadOptions { HasHeaderRow = false });

            Assert.True(reader.Read());
            Assert.Equal(double.MaxValue, Assert.IsType<double>(reader.GetValue(0)));
            Assert.Equal(double.MaxValue, reader.GetDouble(0));
            Assert.False(reader.Read());
        }

        [Fact]
        public void OpenDataReader_XlsFastPathRejectsTruncatedMulBlankRecords() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateTruncatedMulBlankWorkbookStream());

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.OpenDataReader(
                    compound,
                    new ExcelReadOptions { HasHeaderRow = false }));

            Assert.Contains("MULBLANK", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void OpenDataReader_XlsFastPathBoundsDeclaredRootMiniStreamBeforeBuildingItsChain() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateMiniStreamWorkbookCompoundFileWithDeclaredRootSize(
                LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream(),
                uint.MaxValue);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.OpenDataReader(
                    compound,
                    new ExcelReadOptions { HasHeaderRow = false }));

            Assert.Contains(
                "mini stream exceeds configured or physical bounds",
                exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OpenDataReader_XlsFastPathRejectsMiniSectorBeyondDeclaredRootLength() {
            byte[] compound = LegacyXlsCompoundTestBuilder
                .CreateMiniStreamWorkbookCompoundFileWithWorkbookBeyondDeclaredRoot(
                    LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream());

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.OpenDataReader(
                    compound,
                    new ExcelReadOptions { HasHeaderRow = false }));

            Assert.Contains(
                "declared root mini stream length",
                exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OpenDataReader_XlsFastPathBoundsDeclaredMiniFatBeforeReadingItsChain() {
            byte[] compound = LegacyXlsCompoundTestBuilder
                .CreateMiniStreamWorkbookCompoundFileWithDeclaredMiniFatSectorCount(
                    LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream(),
                    1024);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.OpenDataReader(
                    compound,
                    new ExcelReadOptions { HasHeaderRow = false }));

            Assert.Contains(
                "allocation table counts exceed configured or physical bounds",
                exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OpenDataReader_XlsFastPathBoundsDeclaredRegularWorkbookBeforeBuildingItsChain() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithDeclaredWorkbookSize(
                LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream(),
                1024 * 1024);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ExcelDocument.OpenDataReader(
                    compound,
                    new ExcelReadOptions { HasHeaderRow = false }));

            Assert.Contains("physical bounds", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OpenDataReader_XlsFastPathPreservesTheCanonicalSchemaContract() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream());

            using DbDataReader reader = ExcelDocument.OpenDataReader(compound);

            global::OfficeIMO.Excel.Tests.DataReaderSchemaContractAssertions.AssertCanonicalSchema(reader);
        }

        [Theory]
        [InlineData("Workbook")]
        [InlineData("Book")]
        public void OpenDataReader_XlsFastPathIgnoresNestedWorkbookNamedStreams(string nestedStreamName) {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithNestedWorkbookStream(
                LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream(),
                nestedStreamName);

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                compound,
                new ExcelReadOptions { HasHeaderRow = false });

            Assert.Equal(3, reader.FieldCount);
            Assert.True(reader.Read());
            Assert.Equal("Inline", reader.GetString(0));
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateMulRkDateHeaderWorkbookStream() =>
                CreateFastPathRegressionWorkbookStream(
                    "MulRkHeader",
                    globals => WriteRecord(globals, 0x00e0, BuildXfPayload(14)),
                    worksheet => {
                        WriteRecord(worksheet, 0x00bd, BuildMulRkPayload(
                            0,
                            0,
                            (0, EncodeRkInteger(45293))));
                        WriteRecord(worksheet, 0x0203, BuildNumberPayload(1, 0, 1D));
                    });

            internal static byte[] CreateInheritedDateStyleWorkbookStream() =>
                CreateFastPathRegressionWorkbookStream(
                    "InheritedDate",
                    globals => {
                        WriteRecord(globals, 0x00e0, BuildXfPayload(14, isStyle: true));
                        WriteRecord(globals, 0x00e0, BuildXfPayload(
                            0,
                            parentStyleIndex: 0,
                            applyNumberFormat: false));
                    },
                    worksheet => WriteRecord(
                        worksheet,
                        0x0203,
                        BuildNumberPayload(0, 0, 45293D, styleIndex: 1)));

            internal static byte[] CreateInvalidDateSerialWorkbookStream() =>
                CreateFastPathRegressionWorkbookStream(
                    "InvalidDate",
                    globals => WriteRecord(globals, 0x00e0, BuildXfPayload(14)),
                    worksheet => WriteRecord(
                        worksheet,
                        0x0203,
                        BuildNumberPayload(0, 0, double.MaxValue)));

            internal static byte[] CreateTruncatedMulBlankWorkbookStream() =>
                CreateFastPathRegressionWorkbookStream(
                    "TruncatedBlank",
                    globals => { },
                    worksheet => {
                        using var payload = new MemoryStream();
                        WriteUInt16(payload, 0);
                        WriteUInt16(payload, 0);
                        WriteUInt16(payload, 2);
                        WriteRecord(worksheet, 0x00be, payload.ToArray());
                    });

            private static byte[] CreateFastPathRegressionWorkbookStream(
                string sheetName,
                Action<Stream> writeGlobals,
                Action<Stream> writeWorksheet) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, sheetName));
                writeGlobals(stream);
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                writeWorksheet(stream);
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(
                    BitConverter.GetBytes(sheetOffset),
                    0,
                    bytes,
                    checked((int)boundSheetPosition + 4),
                    4);
                return bytes;
            }
        }
    }
}
