using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Theory]
    [InlineData(3)]
    [InlineData(5)]
    public void ExcelHtml_GenericTableWrapsLongLabelsBesidePopulatedColumns(int columnCount) {
        string label = "Known natural satellites as of January 2013; "
            + "the published count includes confirmed observations and retains this explanatory label.";
        string html = "<table><tr><th>Quantity</th>"
            + string.Concat(Enumerable.Range(2, columnCount - 1).Select(index => $"<th>Planet {index}</th>"))
            + "</tr><tr><td><a href='https://example.test/satellites'>" + label + "</a></td>"
            + string.Concat(Enumerable.Range(2, columnCount - 1).Select(index => $"<td>{60 + index}</td>"))
            + "</tr></table>";

        using ExcelDocument workbook = HtmlConversionDocument.Parse(html).ToExcelDocument(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using MemoryStream artifact = workbook.ToStream();
        byte[] bytes = artifact.ToArray();
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(new MemoryStream(bytes), false)) {
            WorksheetPart worksheet = Assert.Single(package.WorkbookPart!.WorksheetParts);
            Column[] columns = worksheet.Worksheet.GetFirstChild<Columns>()?.Elements<Column>().ToArray()
                ?? Array.Empty<Column>();
            Assert.Equal(columnCount, columns.Length);
            Assert.All(columns, column => Assert.InRange(column.Width!.Value, 12D, 60D));
            Assert.True(columns.Sum(column => column.Width!.Value) <= 75.001D);
            Row row = Assert.Single(worksheet.Worksheet.Descendants<Row>(), item => item.RowIndex!.Value == 2U);
            Assert.True(row.Height!.Value > 30D);
            Cell labelCell = Assert.Single(row.Elements<Cell>(), cell => cell.CellReference!.Value == "A2");
            CellFormat format = package.WorkbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!
                .Elements<CellFormat>().ElementAt((int)labelCell.StyleIndex!.Value);
            Assert.True(format.Alignment?.WrapText?.Value);
        }

        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(bytes));
        ExcelSheet sheet = Assert.Single(reopened.Sheets);
        Assert.Equal(label, sheet.CellAt(2, 1).GetValue<string>());
        Assert.Equal("https://example.test/satellites", sheet.GetHyperlinks()["A2"].Target);
        Assert.True(sheet.GetCellStyle(1, 1).Bold);
        for (int column = 2; column <= columnCount; column++) {
            Assert.Equal((60 + column).ToString(), sheet.CellAt(2, column).GetValue<string>());
        }
    }

    [Fact]
    public void ExcelHtml_GenericPresentationPreservesMergedNotesAndEmptySourceRows() {
        const string note = "This note spans all five columns and must remain one editable cell after sizing.";
        const string html = "<table><tr><th>Quantity</th><th>Mercury</th><th>Venus</th><th>Earth</th><th>Mars</th></tr>"
            + "<tr></tr><tr><td>Known natural satellites as of January 2013</td><td>0</td><td>0</td><td>1</td><td>2</td></tr>"
            + "<tr><td colspan='5'>" + note + "</td></tr></table>";
        using ExcelDocument workbook = HtmlConversionDocument.Parse(html).ToExcelDocument(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using MemoryStream artifact = workbook.ToStream();
        using ExcelDocument reopened = ExcelDocument.Load(artifact);
        ExcelSheet sheet = Assert.Single(reopened.Sheets);
        Assert.Equal("A4:E4", Assert.Single(sheet.GetMergedRanges()).A1Range);
        Assert.Equal(note, sheet.CellAt(4, 1).GetValue<string>());
        Assert.Equal("Known natural satellites as of January 2013", sheet.CellAt(3, 1).GetValue<string>());
        Assert.Equal("2", sheet.CellAt(3, 5).GetValue<string>());
        Assert.DoesNotContain(sheet.EnumerateCells(), cell => cell.Row == 2);
        Assert.True(sheet.GetCellStyle(3, 1).WrapText);
        Assert.True(sheet.GetCellStyle(4, 1).WrapText);
    }
}
