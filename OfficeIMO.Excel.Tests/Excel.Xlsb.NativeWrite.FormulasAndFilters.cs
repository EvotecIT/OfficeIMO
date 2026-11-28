using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Array constants, future functions and AutoFilter criteria that native XLSB saves encode and loads project.</summary>
public sealed class ExcelXlsbFormulasAndFiltersTests {
    [Theory]
    [InlineData("SUM(A1:A3*{1;2;3})", null)]
    [InlineData("SUMPRODUCT(A1:A3,{1;2;3})", "14")]
    [InlineData("SUM({1,2;3,4})+A1", "11")]
    [InlineData("INDEX({\"a\",\"b\";TRUE,#N/A},1,2)", "b")]
    [InlineData("IF(A1>0,SUM({1.5,-2}),{\"x\"})", "-0.5")]
    public void Array_constants_in_cell_formulas_survive_native_xlsb_save(string formula, string? value) {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = AddNumbers(document);
            sheet.Cell(1, 2).SetFormula(formula);
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));

        ExcelSheet data = reloaded["Data"];
        Assert.Equal(formula, data.GetFormulaText(1, 2));
        if (value == null) return; // OfficeIMO does not evaluate implicit array arithmetic.
        data.RecalculateSupportedFormulas();
        Assert.True(data.TryGetCellText(1, 2, out string actual));
        Assert.Equal(value, actual);
    }

    [Fact]
    public void Array_constant_and_future_function_defined_names_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = AddNumbers(document);
            AddDefinedName(document, "Weights", "{1;2;3}");
            AddDefinedName(document, "Spread", "_xlfn.STDEV.S('Data'!$A$1:$A$3)");
            sheet.Cell(1, 2).SetFormula("SUMPRODUCT(A1:A3,Weights)");
            sheet.Cell(2, 2).SetFormula("Spread*2");
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));

        DefinedName[] names = reloaded.WorkbookRoot.DefinedNames!.Elements<DefinedName>().ToArray();
        Assert.Equal("{1;2;3}", Assert.Single(names, name => name.Name?.Value == "Weights").Text);
        Assert.Equal("_xlfn.STDEV.S('Data'!$A$1:$A$3)", Assert.Single(names, name => name.Name?.Value == "Spread").Text);
        // The future function is declared the way SpreadsheetML spells it: a hidden #NAME? name.
        DefinedName future = Assert.Single(names, name => name.Name?.Value == "_xlfn.STDEV.S");
        Assert.Equal("#NAME?", future.Text);
        Assert.True(future.Hidden?.Value);
        ExcelSheet data = reloaded["Data"];
        Assert.Equal("SUMPRODUCT(A1:A3,Weights)", data.GetFormulaText(1, 2));
        data.RecalculateSupportedFormulas();
        Assert.True(data.TryGetCellText(1, 2, out string weighted));
        Assert.Equal("14", weighted);
    }

    [Fact]
    public void Future_function_defined_names_survive_a_conversion_through_xlsx() {
        byte[] native = SaveAsXlsb(CreateXlsx(document => {
            ExcelSheet sheet = AddNumbers(document);
            sheet.Cell(1, 2).SetFormula("STDEV.S(A1:A3)");
        }));

        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(native));
        byte[] projected = loaded.ToBytes(ExcelFileFormat.Xlsx);
        using ExcelDocument roundTripped = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(projected)));

        Assert.Equal("_xlfn.STDEV.S(A1:A3)", roundTripped["Data"].GetFormulaText(1, 2));
        DefinedName name = Assert.Single(roundTripped.WorkbookRoot.DefinedNames!.Elements<DefinedName>(),
            item => item.Name?.Value == "_xlfn.STDEV.S");
        Assert.Equal("#NAME?", name.Text);
    }

    [Theory]
    [InlineData("SUM('Data'!$A$1:$A$3)")]
    [InlineData("'Data'!$A$1*2")]
    [InlineData("ABS('Data'!$A$2)")]
    public void Reference_led_defined_name_expressions_survive_native_xlsb_save(string text) {
        byte[] xlsx = CreateXlsx(document => {
            AddNumbers(document);
            AddDefinedName(document, "Total", text);
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));

        DefinedName name = Assert.Single(reloaded.WorkbookRoot.DefinedNames!.Elements<DefinedName>(), item => item.Name?.Value == "Total");
        Assert.Equal(text, name.Text);
    }

    [Theory]
    [InlineData("'Notes _xlfn.Data'!A1*2")]
    [InlineData("IF(A1>0,\"a,b)\",\"_xlfn.X(\")")]
    [InlineData("_xlfn.IFNA(MATCH(\"x,(y\",'Notes _xlfn.Data'!A1:A3,0),-1)")]
    [InlineData("_xlfn.CONCAT(\"_xlpm.x\",'Notes _xlfn.Data'!A1)")]
    public void Quoted_text_does_not_change_how_native_xlsb_formulas_are_split(string formula) {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = AddNumbers(document);
            // A quoted sheet name that contains the future-function prefix.
            document.AddSheet("Notes _xlfn.Data").Cell(1, 1).SetValue(5);
            sheet.Cell(1, 2).SetFormula(formula);
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));

        Assert.Equal(formula, reloaded["Data"].GetFormulaText(1, 2));
    }

    [Theory]
    [InlineData("STDEV.S(A1:A3)", "_xlfn.STDEV.S(A1:A3)", "1")]
    [InlineData("1+CONCAT(\"a\",A2)", "1+_xlfn.CONCAT(\"a\",A2)", null)]
    [InlineData("_xlfn.CONCAT (\"a\",A2)", "_xlfn.CONCAT(\"a\",A2)", null)]
    [InlineData("IFNA(MATCH(9,A1:A3,0),-1)*2", "_xlfn.IFNA(MATCH(9,A1:A3,0),-1)*2", "-2")]
    [InlineData("SUM(A1,MAXIFS(A1:A3,A1:A3,\"<3\"))", "SUM(A1,_xlfn.MAXIFS(A1:A3,A1:A3,\"<3\"))", "3")]
    public void Future_function_cell_formulas_survive_native_xlsb_save(string authored, string stored, string? value) {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = AddNumbers(document);
            sheet.Cell(1, 2).SetFormula(authored);
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));

        ExcelSheet data = reloaded["Data"];
        Assert.Equal(stored, data.GetFormulaText(1, 2));
        string function = stored.Substring(stored.IndexOf("_xlfn.", StringComparison.Ordinal)).Split('(')[0];
        Assert.Contains(reloaded.WorkbookRoot.DefinedNames!.Elements<DefinedName>(),
            name => name.Name?.Value == function && name.Text == "#NAME?" && name.Hidden?.Value == true);
        if (value == null) return;
        data.RecalculateSupportedFormulas();
        Assert.True(data.TryGetCellText(1, 2, out string actual));
        Assert.Equal(value, actual);
    }

    [Fact]
    public void Editing_formulas_of_a_loaded_xlsb_reuses_declared_future_functions_and_refuses_new_ones() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => {
            ExcelSheet sheet = AddNumbers(document);
            sheet.Cell(1, 2).SetFormula("STDEV.S(A1:A3)");
            sheet.Cell(2, 2).SetFormula("SUM({1,2,3})");
        }));

        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].Cell(3, 2).SetFormula("STDEV.S(A1:A2)+SUM({4;5})");
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }
        // Names cannot change on a rewrite: the workbook part is copied as it was.
        Assert.Equal(ReadPart(source, "xl/workbook.bin"), ReadPart(edited, "xl/workbook.bin"));
        using (ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited))) {
            ExcelSheet data = reloaded["Data"];
            Assert.Equal("_xlfn.STDEV.S(A1:A3)", data.GetFormulaText(1, 2));
            Assert.Equal("SUM({1,2,3})", data.GetFormulaText(2, 2));
            Assert.Equal("_xlfn.STDEV.S(A1:A2)+SUM({4;5})", data.GetFormulaText(3, 2));
        }

        using ExcelDocument refused = ExcelDocument.Load(new MemoryStream(source));
        refused["Data"].Cell(3, 2).SetFormula("CONCAT(A1,A2)");
        NotSupportedException exception = Assert.Throws<NotSupportedException>(
            () => refused.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        Assert.Contains("_xlfn.CONCAT", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("contains", "<x:customFilters xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:customFilter operator=\"equal\" val=\"*pe*\" /></x:customFilters>")]
    [InlineData("between", "<x:customFilters and=\"1\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:customFilter operator=\"greaterThanOrEqual\" val=\"2\" /><x:customFilter operator=\"lessThanOrEqual\" val=\"2.5\" /></x:customFilters>")]
    [InlineData("notBetween", "<x:customFilters xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:customFilter operator=\"lessThan\" val=\"2\" /><x:customFilter operator=\"greaterThan\" val=\"3\" /></x:customFilters>")]
    [InlineData("bottomPercent", "<x:top10 top=\"0\" percent=\"1\" val=\"40\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />")]
    [InlineData("topApplied", "<x:top10 top=\"1\" percent=\"0\" val=\"2\" filterVal=\"2\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />")]
    [InlineData("aboveAverage", "<x:dynamicFilter type=\"aboveAverage\" val=\"2\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />")]
    [InlineData("thisMonth", "<x:dynamicFilter type=\"thisMonth\" val=\"45962\" maxVal=\"45992\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />")]
    [InlineData("march", "<x:dynamicFilter type=\"M3\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />")]
    [InlineData("icon", "<x:iconFilter iconSet=\"4Rating\" iconId=\"3\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />")]
    public void AutoFilter_criteria_survive_native_xlsb_save(string criterion, string expected) {
        byte[] xlsx = CreateXlsx(document => AddFilteredTable(document, criterion));

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));

        AutoFilter filter = reloaded["Data"].WorksheetPart.Worksheet!.GetFirstChild<AutoFilter>()!;
        Assert.Equal("A1:B4", filter.Reference?.Value);
        FilterColumn column = Assert.Single(filter.Elements<FilterColumn>());
        Assert.Equal(1U, column.ColumnId?.Value);
        Assert.Equal(expected, Assert.Single(column.ChildElements).OuterXml);
    }

    [Fact]
    public void Editing_a_loaded_xlsb_copies_unchanged_filter_criteria_and_reencodes_edited_ones() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => AddFilteredTable(document, "between")));

        byte[] cellEdit;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].Cell(4, 2).SetValue(7);
            cellEdit = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }
        Assert.Equal(ReadRecords(source, 161, 162), ReadRecords(cellEdit, 161, 162));

        byte[] filterEdit;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(cellEdit))) {
            loaded["Data"].Range("A1:B4").AutoFilter().Column(1U).TopBottom(2);
            filterEdit = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(filterEdit));
        FilterColumn column = Assert.Single(reloaded["Data"].WorksheetPart.Worksheet!.GetFirstChild<AutoFilter>()!.Elements<FilterColumn>());
        Top10 top10 = Assert.IsType<Top10>(Assert.Single(column.ChildElements));
        Assert.True(top10.Top?.Value);
        Assert.False(top10.Percent?.Value);
        Assert.Equal(2D, top10.Val?.Value);
        Assert.True(reloaded["Data"].TryGetCellText(4, 2, out string edited));
        Assert.Equal("7", edited);
    }

    [Fact]
    public void Number_like_text_criteria_stay_native_instead_of_turning_numeric() {
        byte[] generated = SaveAsXlsb(CreateXlsx(document => AddFilteredTable(document, "contains")));
        // Change the stored string comparison "*pe*" into the text "0012"; SpreadsheetML would re-encode it as a number.
        byte[] source = RewriteSheetRecord(generated, 174, data => {
            byte[] copy = (byte[])data.Clone();
            System.Text.Encoding.Unicode.GetBytes("0012").CopyTo(copy, 14);
            return copy;
        });

        byte[] saved;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            AutoFilter filter = loaded["Data"].WorksheetPart.Worksheet!.GetFirstChild<AutoFilter>()!;
            Assert.Empty(filter.Elements<FilterColumn>());
            loaded["Data"].Cell(4, 2).SetValue(7);
            saved = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }
        Assert.Equal(ReadRecords(source, 161, 162), ReadRecords(saved, 161, 162));
    }

    private static byte[] RewriteSheetRecord(byte[] xlsb, int recordType, Func<byte[], byte[]> rewrite) {
        using var output = new MemoryStream();
        using (var source = new ZipArchive(new MemoryStream(xlsb), ZipArchiveMode.Read))
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in source.Entries) {
                byte[] bytes = ReadPart(xlsb, entry.FullName);
                if (entry.FullName == "xl/worksheets/sheet1.bin") {
                    using var rewritten = new MemoryStream();
                    foreach (XlsbRecord record in XlsbRecordReader.ReadAll(new MemoryStream(bytes))) {
                        XlsbRecordWriter.Write(rewritten, record.Type, record.Type == recordType ? rewrite(record.Data) : record.Data);
                    }
                    bytes = rewritten.ToArray();
                }
                using Stream written = target.CreateEntry(entry.FullName).Open();
                written.Write(bytes, 0, bytes.Length);
            }
        }
        return output.ToArray();
    }

    [Fact]
    public void Color_filters_are_refused_because_native_xlsb_styles_have_no_differential_formats() {
        byte[] xlsx = CreateXlsx(document => AddFilteredTable(document, "color"));

        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        NotSupportedException exception = Assert.Throws<NotSupportedException>(
            () => loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        Assert.Contains("differential formats", exception.Message, StringComparison.Ordinal);
    }

    private static ExcelSheet AddNumbers(ExcelDocument document) {
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue(1);
        sheet.Cell(2, 1).SetValue(2);
        sheet.Cell(3, 1).SetValue(3);
        return sheet;
    }

    private static void AddFilteredTable(ExcelDocument document, string criterion) {
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Name");
        sheet.Cell(1, 2).SetValue("Score");
        string[] names = { "Open", "Spec", "Done" };
        for (int row = 2; row <= 4; row++) {
            sheet.Cell(row, 1).SetValue(names[row - 2]);
            sheet.Cell(row, 2).SetValue(row);
        }
        ExcelAutoFilterColumn column = sheet.Range("A1:B4").AutoFilter().Column(1U);
        switch (criterion) {
            case "contains": column.Contains("pe"); return;
            case "between": column.Between(2, 2.5); return;
            case "notBetween": column.NotBetween(2, 3); return;
            case "bottomPercent": column.TopBottom(40, bottom: true, percent: true); return;
        }
        column.TopBottom(1);
        FilterColumn filterColumn = sheet.WorksheetPart.Worksheet!.GetFirstChild<AutoFilter>()!.GetFirstChild<FilterColumn>()!;
        filterColumn.RemoveAllChildren();
        filterColumn.Append(criterion switch {
            "topApplied" => new Top10 { Top = true, Percent = false, Val = 2D, FilterValue = 2D },
            "aboveAverage" => new DynamicFilter { Type = DynamicFilterValues.AboveAverage, Val = 2D },
            "thisMonth" => new DynamicFilter { Type = DynamicFilterValues.ThisMonth, Val = 45962D, MaxVal = 45992D },
            "march" => new DynamicFilter { Type = DynamicFilterValues.March },
            "icon" => new IconFilter { IconSet = IconSetValues.FourRating, IconId = 3U },
            "color" => new ColorFilter { FormatId = 0U, CellColor = true },
            _ => throw new ArgumentOutOfRangeException(nameof(criterion))
        });
    }

    private static void AddDefinedName(ExcelDocument document, string name, string text) {
        DefinedNames names = document.WorkbookRoot.DefinedNames ??= new DefinedNames();
        names.Append(new DefinedName { Name = name, Text = text });
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

    private static byte[] ReadPart(byte[] package, string partName) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var part = new MemoryStream();
        using (Stream entry = archive.GetEntry(partName)!.Open()) entry.CopyTo(part);
        return part.ToArray();
    }

    private static string[] ReadRecords(byte[] xlsb, int beginType, int endType) {
        using var part = new MemoryStream(ReadPart(xlsb, "xl/worksheets/sheet1.bin"));
        var records = XlsbRecordReader.ReadAll(part).ToList();
        int begin = records.FindIndex(record => record.Type == beginType);
        int end = records.FindIndex(record => record.Type == endType);
        Assert.True(begin >= 0 && end > begin);
        return records.Skip(begin).Take(end - begin + 1)
            .Select(record => record.Type + ":" + Convert.ToBase64String(record.Data)).ToArray();
    }
}
