using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Worksheet tables (list objects) that the native XLSB writers encode as Table parts and the reader projects.</summary>
public sealed class ExcelXlsbTableTests {
    public static IEnumerable<object[]> TableShapes() {
        yield return new object[] { "HeaderAndFilter" };
        yield return new object[] { "StyleFlagsWithoutFilter" };
        yield return new object[] { "NoHeaderRow" };
        yield return new object[] { "TotalsRow" };
        yield return new object[] { "TwoTablesOnOneSheet" };
    }

    [Theory]
    [MemberData(nameof(TableShapes))]
    public void Tables_survive_native_xlsb_save(string shape) {
        byte[] xlsx = CreateXlsx(document => BuildShape(document.AddSheet("Data"), shape));
        ExcelTableInfo[] expected;
        using (ExcelDocument source = ExcelDocument.Load(new MemoryStream(xlsx))) expected = source["Data"].GetTables().ToArray();

        byte[] xlsb = SaveAsXlsb(xlsx);

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(xlsb));
        ExcelTableInfo[] actual = reloaded["Data"].GetTables().ToArray();
        Assert.Equal(expected.Length, actual.Length);
        for (int index = 0; index < expected.Length; index++) AssertSameTable(expected[index], actual[index]);
        Assert.Equal(expected.Length, ReadEntryNames(xlsb).Count(name => name.StartsWith("xl/tables/table", StringComparison.Ordinal)));
        if (shape == "TotalsRow") {
            Assert.Equal("SUBTOTAL(109,Sales[Amount])", reloaded["Data"].Cell(5, 3).GetValue().Formula);
            Assert.Equal("Total", reloaded["Data"].Cell(5, 1).GetValue<string>());
        }
    }

    [Fact]
    public void Calculated_columns_and_custom_totals_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            WriteRows(sheet, 1, 1, new[] { "Region", "Amount", "Double" },
                new object[] { "North", 10 }, new object[] { "South", 20 }, new object[] { "East", 30 });
            for (int row = 2; row <= 4; row++) sheet.Cell(row, 3).SetFormula("Sales[@Amount]*2");
            sheet.Cell(5, 1).SetValue("Total");
            sheet.Cell(5, 3).SetFormula("SUM(Sales[Double])/2");
            sheet.AddTable("A1:C5", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2, includeAutoFilter: false);
            // The public table API has no calculated-column or custom-totals setters; state them in the table definition.
            Table table = sheet.WorksheetPart.TableDefinitionParts.Single().Table!;
            table.TotalsRowCount = 1U;
            table.TotalsRowShown = true;
            TableColumn[] columns = table.TableColumns!.Elements<TableColumn>().ToArray();
            columns[0].TotalsRowLabel = "Total";
            columns[2].Append(new CalculatedColumnFormula("Sales[@Amount]*2"));
            columns[2].TotalsRowFunction = TotalsRowFunctionValues.Custom;
            columns[2].Append(new TotalsRowFormula("SUM(Sales[Double])/2"));
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));
        ExcelSheet data = reloaded["Data"];
        TableColumn doubled = data.WorksheetPart.TableDefinitionParts.Single().Table!.TableColumns!.Elements<TableColumn>().Last();
        Assert.Equal("Sales[@Amount]*2", doubled.CalculatedColumnFormula?.Text);
        Assert.Equal(TotalsRowFunctionValues.Custom, doubled.TotalsRowFunction?.Value);
        Assert.Equal("SUM(Sales[Double])/2", doubled.TotalsRowFormula?.Text);
        Assert.Equal("Sales[@Amount]*2", data.Cell(3, 3).GetValue().Formula);
        Assert.Equal("SUM(Sales[Double])/2", data.Cell(5, 3).GetValue().Formula);
        Assert.Equal("Total", data.Cell(5, 1).GetValue<string>());
    }

    [Fact]
    public void Table_local_references_and_array_constants_in_table_formulas_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            WriteRows(sheet, 1, 1, new[] { "Region", "Amount", "Double" },
                new object[] { "North", 10 }, new object[] { "South", 20 }, new object[] { "East", 30 });
            sheet.Cell(5, 1).SetValue("Total");
            sheet.AddTable("A1:C5", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2, includeAutoFilter: false);
            Table table = sheet.WorksheetPart.TableDefinitionParts.Single().Table!;
            table.TotalsRowCount = 1U;
            table.TotalsRowShown = true;
            TableColumn[] columns = table.TableColumns!.Elements<TableColumn>().ToArray();
            columns[0].TotalsRowLabel = "Total";
            // Inside its own table definition a calculated column may leave the table name out.
            columns[2].Append(new CalculatedColumnFormula("[@Amount]*2"));
            columns[2].TotalsRowFunction = TotalsRowFunctionValues.Custom;
            columns[2].Append(new TotalsRowFormula("SUM({1,2})"));
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));

        TableColumn doubled = reloaded["Data"].WorksheetPart.TableDefinitionParts.Single().Table!.TableColumns!.Elements<TableColumn>().Last();
        Assert.Equal("Sales[@Amount]*2", doubled.CalculatedColumnFormula?.Text);
        Assert.Equal("SUM({1,2})", doubled.TotalsRowFormula?.Text);
    }

    [Fact]
    public void Table_top_ten_filters_survive_native_xlsb_save_and_allow_table_edits() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            BuildShape(sheet, "HeaderAndFilter");
            AutoFilter filter = sheet.WorksheetPart.TableDefinitionParts.Single().Table!.AutoFilter!;
            filter.Append(new FilterColumn(new Top10 { Top = true, Percent = false, Val = 2D }) { ColumnId = 2U });
        }));

        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            Top10 top10 = Assert.IsType<Top10>(Assert.Single(loaded["Data"].WorksheetPart.TableDefinitionParts.Single().Table!
                .AutoFilter!.Elements<FilterColumn>()).FirstChild);
            Assert.Equal(2D, top10.Val?.Value);
            loaded["Data"].SetTableStyle("Sales", ExcelTableStyle.TableStyleLight9);
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited));
        Table table = reloaded["Data"].WorksheetPart.TableDefinitionParts.Single().Table!;
        Assert.Equal("TableStyleLight9", table.TableStyleInfo?.Name?.Value);
        Assert.IsType<Top10>(Assert.Single(table.AutoFilter!.Elements<FilterColumn>()).FirstChild);
    }

    [Fact]
    public void Reordering_table_columns_of_a_loaded_xlsb_is_refused() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => BuildShape(document.AddSheet("Data"), "HeaderAndFilter")));
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source));
        TableColumn[] columns = loaded["Data"].WorksheetPart.TableDefinitionParts.Single().Table!.TableColumns!.Elements<TableColumn>().ToArray();
        // Retained formulas address table columns by position, so moving a column id would retarget them.
        (columns[0].Id, columns[1].Id) = (columns[1].Id!.Value, columns[0].Id!.Value);

        NotSupportedException error = Assert.Throws<NotSupportedException>(() => loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        Assert.Contains("reorder columns", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Structured_references_keep_spaces_inside_column_names() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            WriteRows(sheet, 1, 1, new[] { "Amount", " Amount " }, new object[] { 1, 10 }, new object[] { 2, 20 });
            sheet.AddTable("A1:B3", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2, includeAutoFilter: false);
            sheet.Cell(5, 1).SetFormula("SUM(Sales[[ Amount ]])");
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));

        // " Amount " is its own column; trimming it would sum the first column instead.
        Assert.Equal("SUM(Sales[[ Amount ]])", reloaded["Data"].Cell(5, 1).GetValue().Formula);
    }

    [Fact]
    public void A_new_workbook_saved_as_xlsb_with_a_table_and_a_comment_saves_again_after_an_edit() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
        try {
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Data");
                BuildShape(sheet, "HeaderAndFilter");
                // The table takes the first relationship id; the comment parts are added after it.
                sheet.SetComment("E1", "Checked", "Alice");
                document.Save(path);
                sheet.Cell(2, 1).SetValue("West");
                document.Save();
            }

            using ExcelDocument reloaded = ExcelDocument.Load(path);
            Assert.Equal("West", reloaded["Data"].Cell(2, 1).GetValue<string>());
            Assert.Equal("Sales", Assert.Single(reloaded["Data"].GetTables()).Name);
            Assert.Equal("Checked", Assert.Single(reloaded["Data"].GetComments()).Text);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void Editing_cells_of_a_loaded_xlsb_keeps_its_native_table_parts() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => BuildShape(document.AddSheet("Data"), "TwoTablesOnOneSheet")));

        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].Cell(3, 2).SetValue("West");
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        Assert.Equal(ReadEntry(source, "xl/tables/table1.bin"), ReadEntry(edited, "xl/tables/table1.bin"));
        Assert.Equal(ReadEntry(source, "xl/tables/table2.bin"), ReadEntry(edited, "xl/tables/table2.bin"));
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited));
        Assert.Equal("West", reloaded["Data"].Cell(3, 2).GetValue<string>());
        Assert.Equal(new[] { "Sales", "Regions" }, reloaded["Data"].GetTables().Select(table => table.Name).ToArray());
    }

    [Fact]
    public void Restyling_a_table_of_a_loaded_xlsb_rewrites_only_that_table_part() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => BuildShape(document.AddSheet("Data"), "TwoTablesOnOneSheet")));

        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].SetTableStyle("Regions", ExcelTableStyle.TableStyleDark3, showFirstColumn: true, showRowStripes: false);
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        Assert.Equal(ReadEntry(source, "xl/tables/table1.bin"), ReadEntry(edited, "xl/tables/table1.bin"));
        Assert.NotEqual(ReadEntry(source, "xl/tables/table2.bin"), ReadEntry(edited, "xl/tables/table2.bin"));
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited));
        ExcelTableInfo regions = reloaded["Data"].GetTables().Single(table => table.Name == "Regions");
        Assert.Equal("TableStyleDark3", regions.StyleName);
        Assert.True(regions.ShowFirstColumn);
        Assert.False(regions.ShowRowStripes);
        Assert.Equal("TableStyleMedium2", reloaded["Data"].GetTables().Single(table => table.Name == "Sales").StyleName);
    }

    [Fact]
    public void Adding_a_table_to_a_loaded_xlsb_is_refused_instead_of_dropped() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => BuildShape(document.AddSheet("Data"), "HeaderAndFilter")));
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source));
        ExcelSheet sheet = loaded["Data"];
        sheet.Cell(10, 1).SetValue("Code");
        sheet.Cell(11, 1).SetValue("A");
        sheet.AddTable("A10:A11", hasHeader: true, name: "Codes");

        NotSupportedException error = Assert.Throws<NotSupportedException>(() => loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        Assert.Contains("add or remove tables", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Excel_authored_tables_and_structured_references_load_from_xlsb() {
        using ExcelDocument document = ExcelDocument.Load(ExcelFixture("tables.xlsb"));
        ExcelSheet sheet = document["Data"];

        ExcelTableInfo table = Assert.Single(sheet.GetTables());
        Assert.Equal("Sales", table.Name);
        Assert.Equal("A1:C5", table.Range);
        Assert.Equal("TableStyleMedium2", table.StyleName);
        Assert.True(table.HasHeaderRow);
        Assert.True(table.TotalsRowShown);
        Assert.True(table.HasAutoFilter);
        Assert.Equal(new[] { ("Region", (string?)null), ("Product", null), ("Amount", "sum") },
            table.Columns.Select(column => (column.Name, column.TotalsRowFunction)));
        Assert.Equal("SUBTOTAL(109,Sales[Amount])", sheet.Cell(5, 3).GetValue().Formula);
        Assert.Equal(ExcelAuthoredFormulas, ReadFormulas(sheet));
    }

    [Fact]
    public void Editing_an_excel_authored_xlsb_keeps_its_table_part_and_encodes_new_structured_references() {
        byte[] source = File.ReadAllBytes(ExcelFixture("tables.xlsb"));

        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].Cell(2, 1).SetValue("West");
            loaded["Data"].Cell(8, 1).SetFormula("SUM(Sales[[#Totals],[Amount]])+COUNTA(Sales[[#Headers],[Region]])");
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        Assert.Equal(ReadEntry(source, "xl/tables/table1.bin"), ReadEntry(edited, "xl/tables/table1.bin"));
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited));
        ExcelSheet sheet = reloaded["Data"];
        Assert.Equal("West", sheet.Cell(2, 1).GetValue<string>());
        Assert.Equal(ExcelAuthoredFormulas, ReadFormulas(sheet));
        Assert.Equal("SUM(Sales[[#Totals],[Amount]])+COUNTA(Sales[[#Headers],[Region]])", sheet.Cell(8, 1).GetValue().Formula);
        Assert.Equal("A1:C5", Assert.Single(sheet.GetTables()).Range);
    }

    // Formulas Excel wrote in A7:E7 of the fixture, one per PtgList column and row shape.
    private static readonly string[] ExcelAuthoredFormulas = {
        "SUM(Sales[Amount])", "SUM(Sales[[#Totals],[Amount]])", "COUNTA(Sales[#All])", "SUM(Sales[[Product]:[Amount]])", "ROWS(Sales)"
    };

    private static string?[] ReadFormulas(ExcelSheet sheet) =>
        Enumerable.Range(1, 5).Select(column => sheet.Cell(7, column).GetValue().Formula).ToArray();

    private static string ExcelFixture(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Documents", "XlsbCorpus", "excel-generated", name);

    private static void BuildShape(ExcelSheet sheet, string shape) {
        WriteRows(sheet, 1, 1, new[] { "Region", "Product", "Amount" },
            new object[] { "North", "Apples", 10 }, new object[] { "South", "Pears", 20 }, new object[] { "East", "Plums", 30 });
        switch (shape) {
            case "HeaderAndFilter":
                sheet.AddTable("A1:C4", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2, includeAutoFilter: true);
                break;
            case "StyleFlagsWithoutFilter":
                sheet.AddTable("A1:C4", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleLight9, includeAutoFilter: false);
                sheet.SetTableStyle("Sales", ExcelTableStyle.TableStyleLight9, showFirstColumn: true, showLastColumn: true,
                    showRowStripes: false, showColumnStripes: true);
                break;
            case "NoHeaderRow":
                sheet.AddTable("A2:C4", hasHeader: false, name: "Sales", style: ExcelTableStyle.TableStyleMedium9, includeAutoFilter: false);
                break;
            case "TotalsRow":
                sheet.AddTable("A1:C4", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2, includeAutoFilter: true);
                sheet.SetTableTotalsByName("Sales", new Dictionary<string, ExcelTableTotalsFunction> { ["Amount"] = ExcelTableTotalsFunction.Sum });
                break;
            case "TwoTablesOnOneSheet":
                sheet.AddTable("A1:C4", hasHeader: true, name: "Sales", style: ExcelTableStyle.TableStyleMedium2, includeAutoFilter: true);
                WriteRows(sheet, 1, 5, new[] { "Code", "Region name" }, new object[] { "N", "North" }, new object[] { "S", "South" });
                sheet.AddTable("E1:F3", hasHeader: true, name: "Regions", style: ExcelTableStyle.TableStyleLight1, includeAutoFilter: false);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(shape), shape, null);
        }
    }

    private static void WriteRows(ExcelSheet sheet, int row, int column, string[] header, params object[][] rows) {
        for (int index = 0; index < header.Length; index++) sheet.Cell(row, column + index).SetValue(header[index]);
        for (int line = 0; line < rows.Length; line++) {
            for (int index = 0; index < rows[line].Length; index++) sheet.Cell(row + 1 + line, column + index).SetValue(rows[line][index]);
        }
    }

    private static void AssertSameTable(ExcelTableInfo expected, ExcelTableInfo actual) {
        Assert.Equal(expected.Name, actual.Name);
        Assert.Equal(expected.DisplayName, actual.DisplayName);
        Assert.Equal(expected.Range, actual.Range);
        Assert.Equal(expected.StyleName, actual.StyleName);
        Assert.Equal(expected.HasHeaderRow, actual.HasHeaderRow);
        Assert.Equal(expected.TotalsRowShown, actual.TotalsRowShown);
        Assert.Equal(expected.HasAutoFilter, actual.HasAutoFilter);
        Assert.Equal(expected.ShowFirstColumn, actual.ShowFirstColumn);
        Assert.Equal(expected.ShowLastColumn, actual.ShowLastColumn);
        Assert.Equal(expected.ShowRowStripes, actual.ShowRowStripes);
        Assert.Equal(expected.ShowColumnStripes, actual.ShowColumnStripes);
        Assert.Equal(expected.Columns.Select(column => (column.Name, column.TotalsRowFunction)),
            actual.Columns.Select(column => (column.Name, column.TotalsRowFunction)));
    }

    private static byte[] CreateXlsx(Action<ExcelDocument> build) {
        using ExcelDocument document = ExcelDocument.Create();
        build(document);
        return document.ToBytes();
    }

    private static byte[] SaveAsXlsb(byte[] xlsx) {
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        return loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
    }

    private static string[] ReadEntryNames(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return archive.Entries.Select(entry => entry.FullName).ToArray();
    }

    private static byte[] ReadEntry(byte[] package, string name) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using Stream stream = archive.GetEntry(name)!.Open();
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }
}
