using OfficeIMO.Excel;
using System.Data;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(2)]
        public async Task Xlsb_DirectTabularSave_PreservesValuesAfterNeutralViewPreflight(int route) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.PreflightWorkbook();
            var table = new DataTable("Input");
            table.Columns.Add("Label", typeof(string));
            table.Columns.Add("Count", typeof(int));
            table.Rows.Add("Retained label", 42);
            sheet.InsertDataTable(table, includeHeaders: true);

            byte[] bytes;
            if (route == 0) {
                bytes = document.ToBytes(ExcelFileFormat.Xlsb);
            } else {
                using var output = new MemoryStream();
                if (route == 1) document.Save(output, ExcelFileFormat.Xlsb);
                else await document.SaveAsync(output, ExcelFileFormat.Xlsb);
                bytes = output.ToArray();
            }

            Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(bytes);
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Label", reader.GetName(0));
            Assert.Equal("Count", reader.GetName(1));
            Assert.True(reader.Read());
            Assert.Equal("Retained label", reader.GetString(0));
            Assert.Equal(42, reader.GetInt32(1));
            Assert.False(reader.Read());
            Assert.False(reader.NextResult());
        }
    }
}
