using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(ExcelFileFormat.Xlsx, ExcelDateSystem.NineteenHundred)]
        [InlineData(ExcelFileFormat.Xlsx, ExcelDateSystem.NineteenFour)]
        [InlineData(ExcelFileFormat.Xls, ExcelDateSystem.NineteenHundred)]
        [InlineData(ExcelFileFormat.Xls, ExcelDateSystem.NineteenFour)]
        public void DirectDataSet_DateMetadata_PreservesValuesInOtherWriters(
            ExcelFileFormat format, ExcelDateSystem dateSystem) {
            var rows = new[] {
                new DirectDateRecord { Name = "date", When = new DateTime(2026, 10, 9, 6, 30, 0), Value = 42.5 },
            };
            using ExcelDocument document = CreateDirectDateDocument(rows, dateSystem, dataTable: true);
            Assert.True(document.HasDeferredDirectDataSetImport);
            byte[] bytes = document.ToBytes(format);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(bytes);
            Assert.True(reader.Read());
            Assert.Equal(rows[0].Name, reader.GetString(0));
            Assert.Equal(rows[0].When, reader.GetDateTime(1));
            Assert.Equal(rows[0].Value, reader.GetDouble(2));
            Assert.False(reader.Read());
            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(bytes, writable: false));
            Assert.Equal(dateSystem, loaded.DateSystem);
        }
    }
}
