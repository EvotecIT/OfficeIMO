#if !NET8_0_OR_GREATER
using System.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed class XlsbTextCompatibilityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void XlsbText_OrdinaryGettersPreserveValuesOnOlderTargets(bool sharedStrings) {
        using var data = new DataSet();
        var table = data.Tables.Add("Data");
        table.Columns.Add("Text", typeof(string));
        table.Columns.Add("Id", typeof(int));
        string?[] values = ["Żółw 🐢\r\n", string.Empty, null, "following"];
        for (int index = 0; index < values.Length; index++) {
            table.Rows.Add(values[index] == null ? DBNull.Value : values[index], index + 1);
        }
        using ExcelDocument document = ExcelDocument.Create();
        document.InsertDataSet(data, createTables: false, includeAutoFilter: false);
        byte[] package = document.ToBytes(ExcelFileFormat.Xlsb,
            new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings });
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package);
        for (int index = 0; index < values.Length; index++) {
            Assert.True(reader.Read());
            Assert.Equal(index + 1, reader.GetInt32(1));
            if (values[index] == null) {
                Assert.True(reader.IsDBNull(0));
                continue;
            }
            Assert.False(reader.IsDBNull(0));
            Assert.Equal(values[index], reader.GetString(0));
            char[] expected = values[index]!.ToCharArray();
            char[] actual = new char[expected.Length];
            Assert.Equal(expected.Length, reader.GetChars(0, 0, null, 0, 0));
            Assert.Equal(expected.Length, reader.GetChars(0, 0, actual, 0, actual.Length));
            Assert.Equal(expected, actual);
            Assert.Throws<InvalidCastException>(() => reader.GetBytes(0, 0, null, 0, 0));
        }
        Assert.False(reader.Read());
    }
}
#endif
