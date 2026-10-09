using System.Data;
using System.Globalization;
using DBAClientX.Dbf;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.CSV;
using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Dbf.Tests;

public sealed class DbfConversionTests {
    [Theory]
    [InlineData("db3")]
    [InlineData("fp")]
    [InlineData("vfp")]
    public void CsvUsesTheStandardTypedReaderAndPreservesMemoAndBinaryValues(string profile) {
        using var reader = DbfDataReader.Open(DbfReaderTests.Fixture(profile + ".dbf"));
        using var output = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteDataReader(output, reader, new CsvSaveOptions { DateTimeFormat = "O", NullValue = "<null>" });
        Assert.False(reader.IsClosed);
        CsvRow[] rows = CsvDocument.Parse(output.ToString()).AsEnumerable().ToArray();
        Assert.Equal(2, rows.Length);
        Assert.Equal(profile == "vfp" ? "Café" : "Café £", rows[0][0]);
        if (profile == "vfp") {
            Assert.Equal("-123.4567", rows[0][3]);
            Assert.Equal("AP8B", rows[0][7]);
            Assert.Equal("<null>", rows[1][0]);
            Assert.Equal("2020-02-29T12:34:56.7890000", rows[0][4]);
        } else {
            Assert.Equal(1234.50m, decimal.Parse((string)rows[0][1]!, CultureInfo.InvariantCulture));
            Assert.Equal("First memo\r\nSecond line: naïve", rows[0][4]);
            Assert.Equal("<null>", rows[1][1]);
        }
    }

    [Theory]
    [InlineData("db3", false)]
    [InlineData("fp", false)]
    [InlineData("vfp", false)]
    [InlineData("vfp", true)]
    public void ExcelReopensTypedRowsThroughStreamingAndBufferedWriters(string profile, bool buffered) {
        using var reader = DbfDataReader.Open(DbfReaderTests.Fixture(profile + ".dbf"));
        using var output = new MemoryStream();
        ExcelDataSetImportResult result = ExcelDocument.WriteDataReader(output, reader,
            new ExcelTabularWriteOptions { UseSharedStrings = buffered, RequireStreaming = !buffered });
        Assert.Equal(2, result.RowCount);
        Assert.False(reader.IsClosed);
        using (var package = SpreadsheetDocument.Open(output, false)) {
            Assert.Empty(new OpenXmlValidator().Validate(package));
            Cell[] cells = package.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Row>().Skip(1).First().Elements<Cell>().ToArray();
            Assert.Equal(profile == "vfp" ? 42m : 1234.50m, decimal.Parse(cells[1].CellValue!.Text, CultureInfo.InvariantCulture));
        }
        output.Position = 0;
        using var reopened = ExcelDocument.OpenDataReader(output);
        Assert.True(reopened.Read());
        Assert.Equal(profile == "vfp" ? "Café" : "Café £", reopened.GetString(0));
        if (profile == "vfp") {
            Assert.Equal("AP8B", reopened.GetString(7));
            Assert.Equal(Convert.ToBase64String(new byte[] { 0, 255, 1, 32, 65, 66, 67, 32 }), reopened.GetString(6));
            Assert.Equal(-123.4567, Convert.ToDouble(reopened.GetValue(3), CultureInfo.InvariantCulture));
        } else {
            Assert.Equal("First memo\r\nSecond line: naïve", reopened.GetString(4));
            Assert.Equal(1234.50, Convert.ToDouble(reopened.GetValue(1), CultureInfo.InvariantCulture));
        }
        Assert.True(reopened.Read());
        Assert.False(reopened.Read());
    }

    [Fact]
    public void GenericBinaryConsumersUseBase64InsteadOfTheClrTypeName() {
        var table = new DataTable("Data");
        table.Columns.Add("Payload", typeof(byte[]));
        table.Rows.Add(new object[] { new byte[] { 0, 255, 1 } });
        using var csvReader = table.CreateDataReader();
        using var text = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteDataReader(text, csvReader);
        Assert.Equal("AP8B", CsvDocument.Parse(text.ToString()).AsEnumerable().Single()[0]);
        using var output = new MemoryStream();
        var dataSet = new DataSet();
        dataSet.Tables.Add(table);
        ExcelDocument.WriteDataSet(output, dataSet);
        output.Position = 0;
        using var reopened = ExcelDocument.OpenDataReader(output);
        Assert.True(reopened.Read());
        Assert.Equal("AP8B", reopened.GetString(0));
    }
}
