using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLongSharedStringsTests {
    [Fact]
    public void LongStringsWithMatchingProbesRemainDistinctAndShareActualDuplicates() {
        var table = new DataTable("Data");
        table.Columns.Add("Notes", typeof(string));
        string[] distinct = new string[600];
        for (int index = 0; index < distinct.Length; index++) {
            distinct[index] = new string('q', 32)
                + (char)('A' + index / 26) + (char)('A' + index % 26)
                + new string('r', 4070);
            table.Rows.Add(distinct[index]);
        }
        string repeated = new string('z', 4104);
        for (int index = 0; index < 200; index++) table.Rows.Add(repeated);

        using var output = new MemoryStream();
        using (var source = table.CreateDataReader()) ExcelDocument.WriteDataReader(output, source);
        byte[] package = output.ToArray();

        using (var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false)) {
            SharedStringTablePart? shared = spreadsheet.WorkbookPart!.SharedStringTablePart;
            Assert.NotNull(shared);
            SharedStringTable sharedTable = shared!.SharedStringTable!;
            Assert.Equal(201U, sharedTable.Count!.Value);
            Assert.Equal(2U, sharedTable.UniqueCount!.Value);
            Assert.Equal(new[] { "Notes", repeated }, sharedTable.Elements<SharedStringItem>().Select(item => item.InnerText));
            WorksheetPart sheet = spreadsheet.WorkbookPart.WorksheetParts.Single();
            Cell[] repeatedCells = sheet.Worksheet.Descendants<Cell>()
                .Where(cell => cell.CellReference is not null
                    && int.TryParse(cell.CellReference.Value!.Substring(1), out int row)
                    && row >= 602 && row <= 801)
                .ToArray();
            Assert.Equal(200, repeatedCells.Length);
            Assert.All(repeatedCells, cell => {
                Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
                Assert.Equal("1", cell.CellValue!.Text);
            });
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        using var reader = ExcelDocumentReader.Open(new MemoryStream(package, writable: false));
        object?[,] values = reader.GetSheet("Data").ReadRange("A1:A801");
        Assert.Equal("Notes", values[0, 0]);
        for (int index = 0; index < distinct.Length; index++) Assert.Equal(distinct[index], values[index + 1, 0]);
        for (int index = 601; index <= 800; index++) Assert.Equal(repeated, values[index, 0]);
    }
}
