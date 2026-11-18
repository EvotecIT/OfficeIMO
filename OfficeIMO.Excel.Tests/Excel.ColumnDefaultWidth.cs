using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    /// <summary>
    /// Excel reads a column definition without a width as a zero-width, hidden column. Columns created only for
    /// visibility, an outline level or a summary marker keep the worksheet's default width in every saved format.
    /// </summary>
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    [InlineData(".xls")]
    public void ColumnsCreatedWithoutAWidth_KeepTheDefaultWidthInEveryFormat(string extension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.ColumnDefaultWidth.{Guid.NewGuid():N}{extension}");
        try {
            double expected;
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Value");
                sheet.SetColumnHidden(2, true);
                sheet.SetColumnHidden(2, false);
                sheet.SetColumnOutline(3, 1);
                sheet.GroupColumns(5, 6, collapsed: true);
                sheet.ClearColumnGroup(5, 6);
                expected = sheet.GetDefaultStoredColumnWidth();
                document.Save();
            }

            using var loaded = ExcelDocument.Load(path);
            Column[] columns = loaded.Sheets[0].WorksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
            Assert.Equal(new uint[] { 2, 3 }, columns.Select(column => column.Min!.Value).ToArray());
            Assert.All(columns, column => {
                Assert.Equal(expected, column.Width?.Value ?? 0D, 3);
                Assert.NotEqual(true, column.Hidden?.Value);
            });
        } finally {
            File.Delete(path);
        }
    }
}
