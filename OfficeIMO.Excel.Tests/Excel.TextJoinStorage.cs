using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelTextJoinStorageTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextJoin_storage_prefixes_nested_calls_and_preserves_literal_and_qualified_text(bool materialized) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.CellValue(1, 1, "a");
        sheet.CellValue(2, 1, "b");
        document.AddWorksheet("O'Brien, TEXTJOIN(").CellValue(1, 1, "quoted");
        if (materialized) sheet.CellAt(1, 1).GetValue();
        const string authored = "IF(TRUE,TEXTJOIN(\",\",TRUE,A1:A2,\"TEXTJOIN(\",'O''Brien, TEXTJOIN('!A1),_xlfn.TEXTJOIN(\"\",FALSE,\"x\"))";
        const string stored = "IF(TRUE,_xlfn.TEXTJOIN(\",\",TRUE,A1:A2,\"TEXTJOIN(\",'O''Brien, TEXTJOIN('!A1),_xlfn.TEXTJOIN(\"\",FALSE,\"x\"))";
        sheet.CellFormula(1, 2, authored);
        using var output = new MemoryStream(); document.Save(output); output.Position = 0;
        using (var archive = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true)) {
            using var part = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            Assert.Equal(stored, XDocument.Load(part).Descendants(ns + "f").Single().Value);
        }
        output.Position = 0;
        using var reopened = ExcelDocument.Load(output);
        Assert.Equal(stored, reopened.Sheets[0].GetFormulaText(1, 2));
        Assert.Equal(1, reopened.Calculate());
        Assert.Equal("a,b,TEXTJOIN(,quoted", reopened.Sheets[0].CellAt(1, 2).GetValue<string>());
    }

    [Theory]
    [InlineData("TEXTJOIN(\",\",FALSE,A1:A2)")]
    [InlineData("TEXTJOIN(\",\",FALSE,#N/A)")]
    [InlineData("TEXTJOIN(#N/A,FALSE,\"text\")")]
    [InlineData("TEXTJOIN(\",\",#N/A,\"text\")")]
    [InlineData("CONCAT(A1:A2)")]
    public void Joining_text_propagates_errors_as_typed_errors(string expression) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.CellFormula(1, 1, "IFNA(#N/A,#N/A)");
        sheet.CellValue(2, 1, "tail");
        sheet.CellFormula(1, 2, expression);
        Assert.True(document.Calculate() > 0);
        using var output = new MemoryStream(); document.Save(output); output.Position = 0;
        using var archive = new ZipArchive(output, ZipArchiveMode.Read);
        using Stream xml = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XElement cell = XDocument.Load(xml).Descendants(ns + "c").Single(c => c.Attribute("r")?.Value == "B1");
        Assert.Equal("e", cell.Attribute("t")?.Value);
        Assert.Equal("#N/A", cell.Element(ns + "v")?.Value);
    }

    [Fact]
    public void TextJoin_result_above_the_cell_text_limit_remains_a_typed_error() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.CellValue(1, 1, new string('x', 20000));
        sheet.CellValue(2, 1, new string('y', 20000));
        sheet.CellFormula(1, 2, "TEXTJOIN(\",\",FALSE,A1:A2)");
        Assert.Equal(1, document.Calculate());
        using var output = new MemoryStream(); document.Save(output); output.Position = 0;
        using var archive = new ZipArchive(output, ZipArchiveMode.Read);
        using Stream xml = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XElement cell = XDocument.Load(xml).Descendants(ns + "c").Single(c => c.Attribute("r")?.Value == "B1");
        Assert.Equal("e", cell.Attribute("t")?.Value);
        Assert.Equal("#VALUE!", cell.Element(ns + "v")?.Value);
    }

    [Fact]
    public void TextJoin_prefix_expansion_cannot_truncate_a_valid_formula_at_the_storage_limit() {
        string expression = "TEXTJOIN(\"\",FALSE,\"" + new string('x', 8190 - "TEXTJOIN(\"\",FALSE,\"\")".Length) + "\")";
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        Assert.Throws<ArgumentException>(() => sheet.CellFormula(1, 1, expression));
    }
}
