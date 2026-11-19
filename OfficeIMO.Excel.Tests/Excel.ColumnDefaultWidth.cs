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

    [Fact]
    public void ColumnsWithTheDefaultWidth_FollowALaterDefaultWidthChange() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.GroupColumns(2, 3);
        sheet.SetColumnWidth(4, ExcelColumnWidth.Characters(12));

        sheet.SetDefaultColumnWidth(ExcelColumnWidth.Characters(20));

        Column[] columns = sheet.WorksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
        Assert.Equal(new double?[] { 20, 20, 12 }, columns.Select(column => column.Width?.Value).ToArray());

        sheet.ClearColumnGroup(2, 3);
        Assert.Equal(new uint[] { 4 }, sheet.WorksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>()
            .Select(column => column.Min!.Value).ToArray());
    }

    [Fact]
    public void AuthoredColumnWidths_StayWhenTheyEqualTheDefault() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        double defaultWidth = sheet.GetDefaultStoredColumnWidth();
        // As loaded from a file: a non-custom width that happens to equal the default, and a phonetic column.
        sheet.WorksheetPart.Worksheet.InsertAt(new Columns(
            new Column { Min = 2, Max = 2, Width = defaultWidth, CustomWidth = false, OutlineLevel = 1 }), 0);
        sheet.GroupColumns(3, 3);
        sheet.WorksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single(column => column.Min!.Value == 3).Phonetic = true;

        sheet.ClearColumnGroup(2, 3);
        sheet.SetDefaultColumnWidth(ExcelColumnWidth.Characters(20));

        Column[] columns = sheet.WorksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
        Assert.Equal(new uint[] { 2, 3 }, columns.Select(column => column.Min!.Value).ToArray());
        Assert.Equal(defaultWidth, columns[0].Width!.Value);
        Assert.True(columns[1].Phonetic!.Value);
    }

    [Fact]
    public void CopiedWorksheets_KeepTheirDefaultWidthColumns() {
        using var document = ExcelDocument.Create();
        ExcelSheet original = document.AddSheet("Data");
        original.Cell(1, 1).SetValue("Value");
        original.GroupColumns(2, 3);

        ExcelSheet copy = document.CopyWorksheet(original, "Copy");
        copy.SetDefaultColumnWidth(ExcelColumnWidth.Characters(20));

        Assert.Equal(new[] { (2U, 20D), (3U, 20D) }, copy.WorksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>()
            .Select(column => (column.Min!.Value, column.Width!.Value)).ToArray());
        copy.ClearColumnGroup(2, 3);
        Assert.Null(copy.WorksheetPart.Worksheet.GetFirstChild<Columns>());
    }

    [Fact]
    public void DefaultWidthColumns_FollowFromTheWidthTheLibraryAssigned() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.SetDefaultColumnWidth(ExcelColumnWidth.Characters(12));
        sheet.GroupColumns(2, 3);
        Worksheet worksheet = sheet.WorksheetPart.Worksheet;
        // The default changes outside the library, and the caller resizes one column directly.
        worksheet.GetFirstChild<SheetFormatProperties>()!.DefaultColumnWidth = 16;
        worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single(column => column.Min!.Value == 3).Width = 15;

        sheet.SetDefaultColumnWidth(ExcelColumnWidth.Characters(20));

        Assert.Equal(new double?[] { 20, 15 }, worksheet.GetFirstChild<Columns>()!.Elements<Column>().Select(column => column.Width?.Value).ToArray());
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    [InlineData(".xls")]
    public void DefaultWidthColumns_StayWithinTheBinaryWidthLimit(string extension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.WideDefault.{Guid.NewGuid():N}{extension}");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue(1);
                sheet.SetImportedBaseColumnWidth(255);
                sheet.SetColumnHidden(2, false);
                Assert.Equal(255D, sheet.WorksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single().Width!.Value);
                document.Save();
            }

            using var loaded = ExcelDocument.Load(path);
            Column column = loaded.Sheets[0].WorksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single();
            Assert.Equal(255D, column.Width!.Value);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void DefaultWidthColumns_SurviveATransactionRollback() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Value");
        sheet.GroupColumns(2, 3);

        Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(
            _ => throw new InvalidOperationException("rolled back"), new ExcelMutationPlanOptions(), CancellationToken.None));
        sheet.SetDefaultColumnWidth(ExcelColumnWidth.Characters(20));

        Assert.Equal(new[] { (2U, 20D), (3U, 20D) }, sheet.WorksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>()
            .Select(column => (column.Min!.Value, column.Width!.Value)).ToArray());
        sheet.ClearColumnGroup(2, 3);
        Assert.Null(sheet.WorksheetPart.Worksheet.GetFirstChild<Columns>());
    }
}
