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
                using DbDataReader reader = ExcelDocument.OpenDataReader(
                    container,
                    new ExcelReadOptions { HasHeaderRow = false });

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
        public void OpenDataReader_XlsFastPathReadsAllCachedFormulaKinds() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreatePhase4FormulaWorkbookStream());

            using DbDataReader reader = ExcelDocument.OpenDataReader(
                compound,
                new ExcelReadOptions { HasHeaderRow = false });

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
            using (DbDataReader dateReader = ExcelDocument.OpenDataReader(
                       dateCompound,
                       new ExcelReadOptions { HasHeaderRow = false })) {
                Assert.True(dateReader.Read());
                Assert.Equal(new DateTime(1904, 1, 2), dateReader.GetDateTime(0));
                Assert.False(dateReader.Read());
            }

            byte[] stringsCompound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateContinuedSharedStringWorkbookStream());
            using DbDataReader stringsReader = ExcelDocument.OpenDataReader(
                stringsCompound,
                new ExcelReadOptions { HasHeaderRow = false });
            Assert.True(stringsReader.Read());
            Assert.Equal("First", stringsReader.GetString(0));
            Assert.Equal("Second", stringsReader.GetString(1));
            Assert.False(stringsReader.Read());
        }

        [Fact]
        public void OpenDataReader_XlsFastPathEnforcesSharedStringLimitsBeforeDecoding() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(
                LegacyXlsTestWorkbookBuilder.CreateContinuedSharedStringWorkbookStream());

            Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(
                compound,
                new ExcelReadOptions { MaxSharedStringItems = 1 }));
            Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(
                compound,
                new ExcelReadOptions { MaxSharedStringItemCharacters = 5 }));
            Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(
                compound,
                new ExcelReadOptions { MaxSharedStringCharacters = 8 }));
        }
    }
}
