using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(true, 3, "A1:C4")]
        [InlineData(false, 3, "A1:C3")]
        [InlineData(true, 0, "A1:C1")]
        [InlineData(false, 0, null)]
        public void Xlsb_StreamingWriter_OutputLoadsWithValuesAndDimensions(bool includeHeaders, int rowCount, string? expectedRange) {
            string path = Path.Combine(_directoryWithFiles, $"StreamingRoundTrip-{includeHeaders}-{rowCount}.xlsb");
            var rows = Enumerable.Range(1, rowCount).Select(index => (Name: "Item " + index, Amount: index * 1.5D, Active: index % 2 == 0)).ToArray();
            using (var writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xlsb, path, new ExcelWriterOptions { IncludeHeaders = includeHeaders, CreateTable = false })) {
                writer.WriteTable("Data", rows, new[] { "Name", "Amount", "Active" }, (row, item) => {
                    row.Write(item.Name);
                    row.Write(item.Amount);
                    row.Write(item.Active);
                });
                writer.WriteTable("Second", new[] { 7 }, new[] { "Value" }, (row, item) => row.Write(item));
            }

            using (var document = ExcelDocument.Load(path, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Equal(new[] { "Data", "Second" }, document.Sheets.Select(sheet => sheet.Name));
                ExcelSheet data = document["Data"];
                int firstDataRow = includeHeaders ? 2 : 1;
                if (includeHeaders) {
                    Assert.True(data.TryGetCellText(1, 2, out string header));
                    Assert.Equal("Amount", header);
                }

                for (int index = 0; index < rowCount; index++) {
                    Assert.True(data.TryGetCellText(firstDataRow + index, 1, out string name));
                    Assert.Equal(rows[index].Name, name);
                }

                using var reader = document.CreateReader();
                if (expectedRange == null) {
                    Assert.False(data.TryGetCellText(1, 1, out _));
                } else {
                    Assert.Equal(expectedRange, reader.GetSheet("Data").GetUsedRangeA1());
                }

                if (rowCount > 0) {
                    object?[,] values = reader.GetSheet("Data").ReadRange(expectedRange!);
                    Assert.Equal(rows[rowCount - 1].Amount, Convert.ToDouble(values[values.GetLength(0) - 1, 1]));
                    Assert.Equal(rows[rowCount - 1].Active, values[values.GetLength(0) - 1, 2]);
                }

                Assert.True(document["Second"].TryGetCellText(includeHeaders ? 2 : 1, 1, out string second));
                Assert.Equal("7", second);
            }

            using (var document = ExcelDocument.Load(path)) {
                document["Data"].Cell(10, 5).SetValue("Edited");
                // A bulk append keeps a worksheet dimension, which must save although the source had none.
                document["Data"].SetValues(
                    Enumerable.Range(20, 5).SelectMany(row => Enumerable.Range(1, 4).Select(column => (row, column, (object)(row * 10 + column)))),
                    ExcelExecutionMode.Sequential);
                document.Save();
            }

            using (var document = ExcelDocument.Load(path, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.True(document["Data"].TryGetCellText(10, 5, out string edited));
                Assert.Equal("Edited", edited);
                Assert.True(document["Data"].TryGetCellText(24, 4, out string appended));
                Assert.Equal("244", appended);
                if (rowCount > 0) {
                    Assert.True(document["Data"].TryGetCellText(firstDataRowFor(includeHeaders), 1, out string name));
                    Assert.Equal(rows[0].Name, name);
                }
            }

            File.Delete(path);

            static int firstDataRowFor(bool headers) => headers ? 2 : 1;
        }
    }
}
