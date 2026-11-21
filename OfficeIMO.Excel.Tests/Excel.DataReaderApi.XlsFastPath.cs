using System.Data.Common;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void OpenDataReader_XlsFastPathReadsRegularAndMiniCompoundStreams() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream();
            byte[][] containers = {
                LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream),
                LegacyXlsCompoundTestBuilder.CreateMiniStreamWorkbookCompoundFile(workbookStream)
            };

            foreach (byte[] container in containers) {
                using DbDataReader reader = ExcelDocument.OpenReader(
                    container,
                    new ExcelReaderOptions { HasHeaderRow = false });

                Assert.Equal(3, reader.FieldCount);
                Assert.True(reader.Read());
                Assert.Equal("Inline", reader.GetString(0));
                Assert.True(reader.IsDBNull(1));
                Assert.True(reader.Read());
                Assert.Equal(7, reader.GetInt32(0));
                Assert.Equal(-3, reader.GetInt32(1));
                Assert.Equal(123.45D, reader.GetDouble(2), precision: 10);
                Assert.True(reader.Read());
                Assert.Equal(1, reader.GetInt32(0));
                Assert.Equal(2, reader.GetInt32(1));
                Assert.True(reader.Read());
                Assert.True(reader.IsDBNull(0));
                Assert.True(reader.IsDBNull(1));
                Assert.True(reader.Read());
                Assert.Equal("#DIV/0!", reader.GetString(0));
                Assert.False(reader.Read());
            }
        }

        [Fact]
        public void OpenDataReader_XlsFastPathInfersSchemaAndReplaysSampledRows() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream());

            using DbDataReader reader = ExcelDocument.OpenReader(
                compound,
                new ExcelReaderOptions {
                    HasHeaderRow = false,
                    InferSchema = true,
                    SchemaSampleRows = 5,
                    NumericAsDecimal = true
                });

            Assert.Equal(typeof(object), reader.GetFieldType(0));
            Assert.Equal(typeof(decimal), reader.GetFieldType(1));
            Assert.Equal(typeof(decimal), reader.GetFieldType(2));

            Assert.True(reader.Read());
            Assert.Equal("Inline", reader.GetString(0));
            Assert.True(reader.IsDBNull(1));
            Assert.True(reader.Read());
            Assert.Equal(7, reader.GetInt32(0));
            Assert.Equal(-3, reader.GetInt32(1));
            Assert.Equal(123.45m, reader.GetDecimal(2));
            Assert.True(reader.Read());
            Assert.Equal(1, reader.GetInt32(0));
            Assert.Equal(2, reader.GetInt32(1));
            Assert.True(reader.Read());
            Assert.True(reader.IsDBNull(0));
            Assert.True(reader.IsDBNull(1));
            Assert.True(reader.Read());
            Assert.Equal("#DIV/0!", reader.GetString(0));
            Assert.False(reader.Read());
        }

        [Fact]
        public void OpenDataReader_XlsFastPathBoundsSchemaSampling() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream());

            InvalidOperationException rowLimit = Assert.Throws<InvalidOperationException>(() =>
{ using var reader = ExcelDocument.OpenReader(
                    compound,
                    new ExcelReaderOptions {
                        HasHeaderRow = false,
                        InferSchema = true,
                        SchemaSampleRows = 2,
                        MaxDataReaderSchemaSampleRows = 1
                    }); _ = reader.FieldCount; });
            Assert.Contains("Schema sample row count", rowLimit.Message, StringComparison.Ordinal);

            InvalidOperationException cellLimit = Assert.Throws<InvalidOperationException>(() =>
{ using var reader = ExcelDocument.OpenReader(
                    compound,
                    new ExcelReaderOptions {
                        HasHeaderRow = false,
                        InferSchema = true,
                        SchemaSampleRows = 2,
                        MaxDataReaderBufferedCells = 3
                    }); _ = reader.FieldCount; });
            Assert.Contains("buffer", cellLimit.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OpenDataReader_XlsFastPathReadsAllCachedFormulaKinds() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreatePhase4FormulaWorkbookStream());

            using DbDataReader reader = ExcelDocument.OpenReader(
                compound,
                new ExcelReaderOptions { HasHeaderRow = false });

            Assert.True(reader.Read());
            Assert.Equal(42D, reader.GetDouble(0));
            Assert.True(reader.GetBoolean(1));
            Assert.Equal("Formula text", reader.GetString(2));
            Assert.Equal("#VALUE!", reader.GetString(3));
            Assert.Equal("Continued formula text", reader.GetString(4));
            Assert.True(reader.Read());
            Assert.Equal(42D, reader.GetDouble(2));
            Assert.Equal(42D, reader.GetDouble(3));
            Assert.False(reader.Read());
        }

        [Fact]
        public void OpenDataReader_XlsFastPathReads1904DatesAndContinuedSharedStrings() {
            byte[] dateCompound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateDate1904DateFormattedWorkbookStream());
            using (DbDataReader dateReader = ExcelDocument.OpenReader(
                       dateCompound,
                       new ExcelReaderOptions { HasHeaderRow = false })) {
                Assert.True(dateReader.Read());
                Assert.Equal(new DateTime(1904, 1, 2), dateReader.GetDateTime(0));
                Assert.False(dateReader.Read());
            }

            byte[] stringsCompound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateContinuedSharedStringWorkbookStream());
            using DbDataReader stringsReader = ExcelDocument.OpenReader(
                stringsCompound,
                new ExcelReaderOptions { HasHeaderRow = false });
            Assert.True(stringsReader.Read());
            Assert.Equal("First", stringsReader.GetString(0));
            Assert.Equal("Second", stringsReader.GetString(1));
            Assert.False(stringsReader.Read());
        }

        [Fact]
        public void OpenDataReader_XlsFastPathEnforcesSharedStringLimitsBeforeDecoding() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateContinuedSharedStringWorkbookStream());

            Assert.Throws<InvalidDataException>(() => { using var reader = ExcelDocument.OpenReader(
                compound,
                new ExcelReaderOptions { MaxSharedStringItems = 1 }); _ = reader.FieldCount; });
            Assert.Throws<InvalidDataException>(() => { using var reader = ExcelDocument.OpenReader(
                compound,
                new ExcelReaderOptions { MaxSharedStringItemCharacters = 5 }); _ = reader.FieldCount; });
            Assert.Throws<InvalidDataException>(() => { using var reader = ExcelDocument.OpenReader(
                compound,
                new ExcelReaderOptions { MaxSharedStringCharacters = 8 }); _ = reader.FieldCount; });
        }

        [Fact]
        public void OpenDataReader_XlsFastPathEnforcesContainerInputLimitForPath() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream());
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.XlsInputLimit.{Guid.NewGuid():N}.xls");
            try {
                File.WriteAllBytes(path, compound);

                InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
{ using var reader = ExcelDocument.OpenReader(
                        path,
                        new ExcelReaderOptions { MaxInputBytes = compound.LongLength - 1L }); _ = reader.FieldCount; });

                Assert.Contains("exceeding the configured limit", exception.Message, StringComparison.OrdinalIgnoreCase);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void OpenDataReader_XlsFastPathRejectsMissingHeaderFormulaString() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateStringFormulaWithoutStringRecordWorkbookStream(
                    formulaInFirstRow: true));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
{ using var reader = ExcelDocument.OpenReader(
                    compound,
                    new ExcelReaderOptions { HasHeaderRow = false }); _ = reader.FieldCount; });

            Assert.Contains("required String record", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void OpenDataReader_XlsFastPathRejectsMissingDataFormulaString() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateStringFormulaWithoutStringRecordWorkbookStream(
                    formulaInFirstRow: false));

            using DbDataReader reader = ExcelDocument.OpenReader(
                compound,
                new ExcelReaderOptions { HasHeaderRow = false });
            Assert.True(reader.Read());
            Assert.Throws<InvalidDataException>(() => reader.Read());
        }

        /// <summary>
        /// A reader of one XLS worksheet loads only the workbook globals and that worksheet, and returns
        /// what the reader of every worksheet returns for it.
        /// </summary>
        [Fact]
        public void OpenDataReader_XlsSelectedSheetMatchesTheSameSheetReadWithTheWholeWorkbook() {
            using var output = new MemoryStream();
            using (var writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xls, output,
                       new ExcelWriterOptions { UseSharedStrings = true }, leaveOpen: true)) {
                for (int sheet = 0; sheet < 3; sheet++) {
                    int rows = 200 + sheet * 1500;
                    writer.WriteTable("Data" + (sheet + 1), Enumerable.Range(0, rows), new[] { "Name", "Id", "Amount" },
                        (row, index) => row.Write("Item " + index % 37).Write(sheet * 10_000 + index).Write(index * 0.25));
                }
            }
            byte[] workbook = output.ToArray();

            var expected = new List<List<object>>();
            using (DbDataReader all = ExcelDocument.OpenReader(workbook)) {
                do {
                    var values = new List<object>();
                    while (all.Read()) {
                        for (int column = 0; column < all.FieldCount; column++) values.Add(all.GetValue(column));
                    }
                    expected.Add(values);
                } while (all.NextResult());
            }

            Assert.Equal(3, expected.Count);
            for (int sheet = 0; sheet < 3; sheet++) {
                foreach (var options in new[] {
                             new ExcelReaderOptions { SheetIndex = sheet },
                             new ExcelReaderOptions { SheetName = "Data" + (sheet + 1) } }) {
                    using DbDataReader single = ExcelDocument.OpenReader(workbook, options);
                    var values = new List<object>();
                    while (single.Read()) {
                        for (int column = 0; column < single.FieldCount; column++) values.Add(single.GetValue(column));
                    }
                    Assert.Equal(expected[sheet], values);
                    Assert.False(single.NextResult());
                }
            }
        }
    }
}
