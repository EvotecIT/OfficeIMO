using OfficeIMO.Excel;
using System.Data;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_DirectTabularSave_RejectsWideDataBeforeChangingDestination() {
            using var data = new DataTable();
            for (int column = 0; column < 257; column++) data.Columns.Add("Column" + column, typeof(int));
            data.Rows.Add(Enumerable.Repeat<object>(1, 257).ToArray());
            using ExcelDocument document = ExcelDocument.Create();
            document.AddSheet("Data").WriteData(data);
            byte[] sentinel = { 1, 2, 3, 4, 5 };
            using var destination = new MemoryStream();
            destination.Write(sentinel);
            destination.Position = 2;

            Assert.Throws<NotSupportedException>(() =>
                document.Save(destination, new ExcelSaveOptions { Format = ExcelFileFormat.Xls }));

            Assert.Equal(sentinel, destination.ToArray());
            Assert.Equal(2, destination.Position);
        }
    }
}
