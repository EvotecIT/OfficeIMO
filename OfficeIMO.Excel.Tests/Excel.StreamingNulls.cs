using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StreamingTableAndReader_DistinguishDatabaseNullFromEmptyText(bool readerSource) {
        using var table = new DataTable();
        table.Columns.Add("Value", typeof(string));
        table.Rows.Add(DBNull.Value);
        table.Rows.Add(string.Empty);
        table.Rows.Add("Łódź");
        using var source = table.CreateDataReader();
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(output, leaveOpen: true)) {
            if (readerSource) writer.WriteTable("Data", source);
            else writer.WriteTable("Data", table);
        }
        Assert.False(source.IsClosed);
        using (var package = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false)) {
            var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Row>()
                .Skip(1).Select(row => row.Elements<Cell>().Single()).ToArray();
            Assert.Null(cells[0].DataType);
            Assert.Null(cells[0].CellValue);
            Assert.NotNull(cells[1].DataType);
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
        using var reader = ExcelDocument.OpenReader(new MemoryStream(output.ToArray()));
        Assert.True(reader.Read());
        Assert.True(reader.IsDBNull(0));
        Assert.True(reader.Read());
        Assert.False(reader.IsDBNull(0));
        Assert.Equal(string.Empty, reader.GetString(0));
        Assert.True(reader.Read());
        Assert.Equal("Łódź", reader.GetString(0));
        Assert.False(reader.Read());
    }
}
