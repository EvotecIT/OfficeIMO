using System;
using System.Data;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelWorksheetUnicodeNamesTests {
    [Theory]
    [InlineData(30)]
    [InlineData(26)]
    public void Sanitized_sheet_names_and_collision_suffixes_save_reopen_and_resolve(int prefixLength) {
        string requested = new string('A', prefixLength) + "😀common";
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet first = document.AddWorksheet(requested);
        first.CellValue(1, 1, "First");
        document.AddWorksheet(requested).CellValue(1, 1, "Second");
        Assert.Equal("First", document.GetOrCreateSheet(requested, ExcelSheetNameValidationMode.Sanitize)
            .CellAt(1, 1).GetValue<string>());
        using var saved = new MemoryStream();
        document.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);

        Assert.Equal(2, reopened.Sheets.Count);
        Assert.NotEqual(reopened.Sheets[0].Name, reopened.Sheets[1].Name);
        Assert.All(reopened.Sheets, sheet => System.Xml.XmlConvert.VerifyXmlChars(sheet.Name));
        Assert.Equal("Second", reopened.Sheets[1].CellAt(1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(30)]
    [InlineData(26)]
    public void Deferred_dataset_export_preserves_unicode_at_both_name_boundaries(int prefixLength) {
        using var data = new DataSet();
        string prefix = new string('A', prefixLength) + "😀common";
        for (int index = 0; index < 2; index++) {
            var table = new DataTable(prefix + index);
            table.Columns.Add("Value", typeof(int));
            table.Rows.Add(index + 1);
            data.Tables.Add(table);
        }
        using ExcelDocument document = ExcelDocument.Create();
        document.InsertDataSet(data, createTables: false);
        using var saved = new MemoryStream();
        document.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);

        Assert.Equal(2, reopened.Sheets.Count);
        Assert.NotEqual(reopened.Sheets[0].Name, reopened.Sheets[1].Name);
        Assert.All(reopened.Sheets, sheet => System.Xml.XmlConvert.VerifyXmlChars(sheet.Name));
        Assert.Equal(1, reopened.Sheets[0].CellAt(2, 1).GetValue<int>());
        Assert.Equal(2, reopened.Sheets[1].CellAt(2, 1).GetValue<int>());
    }
}
