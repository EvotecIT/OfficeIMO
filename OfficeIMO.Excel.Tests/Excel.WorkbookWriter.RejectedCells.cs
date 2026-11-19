using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelWorkbookWriterRejectedCellTests {
    private static readonly string TooLong = new('x', 32_768);

    [Theory]
    [InlineData(ExcelFileFormat.Xlsx)]
    [InlineData(ExcelFileFormat.Xlsb)]
    [InlineData(ExcelFileFormat.Xls)]
    public void RejectedCellKeepsTheRowPositionAfterASkippedCell(ExcelFileFormat format) {
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(format, output, leaveOpen: true)) {
            writer.WriteTable("Data", new[] { 1 }, new[] { "Skipped", "Text", "Number" }, (row, _) => {
                row.WriteBlank();
                Assert.Throws<ArgumentException>(() => row.Write(TooLong));
                row.Write("ok").Write(5);
            });
        }

        using var reader = ExcelDocument.OpenReader(output.ToArray());
        Assert.True(reader.Read());
        Assert.True(reader.IsDBNull(0));
        Assert.Equal("ok", reader.GetString(1));
        Assert.Equal(5d, reader.GetDouble(2));
        Assert.False(reader.Read());
    }

    [Fact]
    public void RejectedCellKeepsTheSelectedStyleForTheNextCell() {
        var options = new ExcelWriterOptions {
            Styles = new Dictionary<string, ExcelStyleDefinition> { ["Mark"] = new() { Bold = true } }
        };
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(output, options, leaveOpen: true)) {
            writer.WriteTable("Data", new[] { 1, 2 }, new[] { "Skipped", "Text" }, (row, value) => {
                row.WriteBlank().SetNextCellStyle("Mark");
                if (value == 1) Assert.Throws<ArgumentException>(() => row.Write(TooLong));
                row.Write("ok");
            });
        }

        output.Position = 0;
        using var package = SpreadsheetDocument.Open(output, false);
        Cell[] cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToArray();
        Cell rejectedRow = cells.Single(cell => cell.CellReference?.Value == "B2");
        Cell controlRow = cells.Single(cell => cell.CellReference?.Value == "B3");
        Assert.NotNull(controlRow.StyleIndex);
        Assert.NotEqual(0U, controlRow.StyleIndex!.Value);
        Assert.Equal(controlRow.StyleIndex.Value, rejectedRow.StyleIndex?.Value);
    }
}
