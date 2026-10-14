using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelFutureFunctionStorageTests {
    [Fact]
    public void Future_functions_match_native_Excel_before_and_after_saved_operand_edits() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "excel-future-functions-v16.113.3.json")));
        using var document = ExcelDocument.Load(Path.Combine(root, manifest.RootElement.GetProperty("input").GetString()!));
        int count = manifest.RootElement.GetProperty("formulaCells").GetInt32();
        Assert.Equal(count, document.Calculate());
        Verify(document.Sheets[0], edited: false);
        using var original = new MemoryStream(); document.Save(original); VerifySavedKinds(original);
        original.Position = 0;
        using var reopened = ExcelDocument.Load(original);
        Verify(reopened.Sheets[0], edited: false);
        Assert.Empty(reopened.ValidateOpenXml());
        reopened.Sheets[0].CellValue(2, 2, 50d);
        Assert.Equal(count, reopened.Calculate());
        Verify(reopened.Sheets[0], edited: true);
        using var changed = new MemoryStream(); reopened.Save(changed); VerifySavedKinds(changed);
        changed.Position = 0;
        using var final = ExcelDocument.Load(changed);
        Verify(final.Sheets[0], edited: true);
        Assert.Empty(final.ValidateOpenXml());

        void Verify(ExcelSheet sheet, bool edited) {
            foreach (JsonElement cell in manifest.RootElement.GetProperty("cells").EnumerateArray()) {
                int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
                Assert.Equal(cell.GetProperty("formula").GetString(), sheet.GetFormulaText(row, column));
                Assert.True(sheet.TryGetCachedFormulaValue(row, column, out string? value), cell.GetProperty("address").GetString());
                Assert.Equal(cell.GetProperty(edited ? "editedValue" : "value").GetString(), value);
            }
        }

        void VerifySavedKinds(MemoryStream bytes) {
            bytes.Position = 0;
            using var archive = new ZipArchive(bytes, ZipArchiveMode.Read, leaveOpen: true);
            using var part = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var cells = XDocument.Load(part).Descendants(ns + "c").ToDictionary(c => c.Attribute("r")!.Value);
            foreach (JsonElement cell in manifest.RootElement.GetProperty("cells").EnumerateArray()) {
                XElement saved = cells[cell.GetProperty("address").GetString()!];
                Assert.Equal(cell.GetProperty("kind").GetString(), saved.Attribute("t")?.Value ?? "n");
                Assert.Equal(cell.GetProperty("formulaType").GetString(), saved.Element(ns + "f")!.Attribute("t")?.Value);
            }
        }
    }

    [Theory]
    [InlineData("buffered")]
    [InlineData("materialized")]
    [InlineData("array")]
    public void Future_function_storage_preserves_calls_literals_references_and_saved_calculation(string route) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("O'Brien, IFNA(MINIFS(");
        var cases = new[] {
            ("IFNA(NA(),99)", "_xlfn.IFNA(NA(),99)", "99"),
            ("IFNA(1/0,99)", "_xlfn.IFNA(1/0,99)", "#DIV/0!"),
            ("MINIFS(B2:B4,A2:A4,\"EU\")", "_xlfn.MINIFS(B2:B4,A2:A4,\"EU\")", "10"),
            ("MAXIFS(B2:B4,A2:A4,\"EU\")", "_xlfn.MAXIFS(B2:B4,A2:A4,\"EU\")", "30"),
            ("MINIFS(B2:B4,A2:A4,\"missing\")", "_xlfn.MINIFS(B2:B4,A2:A4,\"missing\")", "0"),
            ("MAXIFS(B2:B4,A2:A4,\"missing\")", "_xlfn.MAXIFS(B2:B4,A2:A4,\"missing\")", "0"),
            ("ifna(maxifs(B2:B4,A2:A4,\"EU\"),99)", "_xlfn.ifna(_xlfn.maxifs(B2:B4,A2:A4,\"EU\"),99)", "30"),
            ("IFNA(NA(),\"IFNA( \"\"MINIFS(\"\" MAXIFS(\")&'O''Brien, IFNA(MINIFS('!A1",
             "_xlfn.IFNA(NA(),\"IFNA( \"\"MINIFS(\"\" MAXIFS(\")&'O''Brien, IFNA(MINIFS('!A1", "IFNA( \"MINIFS(\" MAXIFS(quoted"),
            ("_xlfn.IFNA(NA(),88)", "_xlfn.IFNA(NA(),88)", "88")
        };
        const int firstFormulaRow = 30;
        // A complete A1-based rectangle reaches the direct serializer, with formulas beyond its buffering threshold.
        for (int row = 1; row < firstFormulaRow + cases.Length; row++) {
            for (int column = 1; column <= 5; column++) {
                if (column == 5 && row >= firstFormulaRow) {
                    string formula = cases[row - firstFormulaRow].Item1;
                    if (route == "array") sheet.SetArrayFormula($"E{row}:E{row}", formula);
                    else sheet.CellFormula(row, column, formula);
                } else {
                    object value = (row, column) switch {
                        (1, 1) => "quoted",
                        (2, 1) or (3, 1) => "EU",
                        (4, 1) => "US",
                        (2, 2) => 10d,
                        (3, 2) => 30d,
                        (4, 2) => 20d,
                        _ => 0d
                    };
                    sheet.CellValue(row, column, value);
                }
            }
        }
        if (route == "materialized") sheet.CellAt(1, 1).GetValue();
        using var output = new MemoryStream(); document.Save(output); output.Position = 0;
        if (route == "buffered") Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true)) {
            using var part = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XElement[] formulas = XDocument.Load(part).Descendants(ns + "f").ToArray();
            Assert.Equal(cases.Length, formulas.Length);
            for (int index = 0; index < cases.Length; index++) {
                Assert.Equal(cases[index].Item2, formulas[index].Value);
                Assert.Equal(route == "array" ? "array" : null, formulas[index].Attribute("t")?.Value);
            }
        }
        output.Position = 0;
        using var reopened = ExcelDocument.Load(output);
        Assert.Equal(cases.Length, reopened.Calculate());
        for (int index = 0; index < cases.Length; index++) {
            Assert.Equal(cases[index].Item2, reopened.Sheets[0].GetFormulaText(index + firstFormulaRow, 5));
            Assert.True(reopened.Sheets[0].TryGetCachedFormulaValue(index + firstFormulaRow, 5, out string? value));
            Assert.Equal(cases[index].Item3, value);
        }
        using var saved = new MemoryStream(); reopened.Save(saved); saved.Position = 0;
        using (var archive = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true)) {
            using var part = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XElement error = XDocument.Load(part).Descendants(ns + "c").Single(c => c.Attribute("r")?.Value == "E31");
            Assert.Equal("e", error.Attribute("t")?.Value);
            Assert.Equal("#DIV/0!", error.Element(ns + "v")?.Value);
        }
        saved.Position = 0;
        using var final = ExcelDocument.Load(saved);
        Assert.Empty(final.ValidateOpenXml());
        for (int index = 0; index < cases.Length; index++) {
            Assert.True(final.Sheets[0].TryGetCachedFormulaValue(index + firstFormulaRow, 5, out string? value));
            Assert.Equal(cases[index].Item3, value);
        }
    }

    [Theory]
    [InlineData("IFNA", false)]
    [InlineData("MINIFS", false)]
    [InlineData("MAXIFS", false)]
    [InlineData("IFNA", true)]
    [InlineData("MINIFS", true)]
    [InlineData("MAXIFS", true)]
    public void Future_function_prefix_expansion_rejects_oversized_formulas_before_mutating_cells(string function, bool array) {
        string prefix = function == "IFNA" ? "IFNA(NA(),\"" : function + "(B1,B1,\"";
        string expression = prefix + new string('x', 8190 - prefix.Length - 2) + "\")";
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.CellValue(1, 1, 17d); sheet.CellFormula(1, 1, "42");
        sheet.CellValue(2, 1, "keep");
        if (array) Assert.Throws<ArgumentException>(() => sheet.SetArrayFormula("A1:A2", expression));
        else Assert.Throws<ArgumentException>(() => sheet.CellFormula(1, 1, expression));
        Assert.Equal("42", sheet.GetFormulaText(1, 1));
        Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? value));
        Assert.Equal("17", value);
        Assert.Equal("keep", sheet.CellAt(2, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Combined_array_and_future_prefixes_enforce_the_final_storage_limit_atomically(bool dynamic) {
        const string prefix = "IFNA(SEQUENCE(1),\"";
        string atLimit = prefix + new string('x', 8180 - prefix.Length - 2) + "\")";
        string oversized = prefix + new string('x', 8181 - prefix.Length - 2) + "\")";
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.CellValue(1, 1, 17d); sheet.CellFormula(1, 1, "42");
        sheet.CellValue(2, 1, "keep");
        Assert.Throws<ArgumentException>(() => Set(oversized));
        Assert.Equal("42", sheet.GetFormulaText(1, 1));
        Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? value));
        Assert.Equal("17", value);
        Assert.Equal("keep", sheet.CellAt(2, 1).GetValue<string>());
        Set(atLimit);
        string stored = sheet.GetFormulaText(1, 1)!;
        Assert.Equal(8192, stored.Length);
        Assert.StartsWith("_xlfn.IFNA(_xlfn.SEQUENCE(1),\"", stored);
        using var output = new MemoryStream(); document.Save(output); output.Position = 0;
        using var reopened = ExcelDocument.Load(output);
        Assert.Equal(stored, reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Empty(reopened.ValidateOpenXml());

        void Set(string formula) {
            if (dynamic) sheet.SetDynamicArrayFormula("A1", formula);
            else sheet.SetArrayFormula("A1:A2", formula);
        }
    }

    [Fact]
    public void Array_storage_qualifies_only_calls_and_preserves_qualified_names_and_literals() {
        const string formula = "IFNA(SEQUENCE (1),SORT(UNIQUE(FILTER(A1:A2,A1:A2>0))))"
            + "&\"SEQUENCE( SORT( IFNA(\"&Report!SEQUENCE(1)&'Report'!IFNA(NA(),99)&Table1[SEQUENCE(]";
        const string stored = "_xlfn.IFNA(_xlfn.SEQUENCE (1),_xlfn._xlws.SORT(_xlfn.UNIQUE(_xlfn._xlws.FILTER(A1:A2,A1:A2>0))))"
            + "&\"SEQUENCE( SORT( IFNA(\"&Report!SEQUENCE(1)&'Report'!IFNA(NA(),99)&Table1[SEQUENCE(]";
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.SetArrayFormula("D1:D1", formula);
        Assert.Equal(stored, sheet.GetFormulaText(1, 4));
        using var output = new MemoryStream(); document.Save(output); output.Position = 0;
        using var reopened = ExcelDocument.Load(output);
        Assert.Equal(stored, reopened.Sheets[0].GetFormulaText(1, 4));
        Assert.Empty(reopened.ValidateOpenXml());
    }
}
