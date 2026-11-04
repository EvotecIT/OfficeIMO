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
    [InlineData("text")]
    [InlineData("inline")]
    [InlineData("string-cache")]
    [InlineData("error-cache")]
    public void Joining_distinguishes_error_like_text_from_typed_formula_errors(string kind) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        if (kind.EndsWith("cache", StringComparison.Ordinal)) sheet.CellFormulaWithTextCache(1, 1, "UNSUPPORTED()", "#N/A");
        else sheet.Cell(1, 1).SetValue("#N/A");
        sheet.Cell(1, 1).GetValue();
        var source = sheet.WorksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single();
        if (kind == "inline") {
            source.CellValue = null;
            source.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
            source.InlineString = new DocumentFormat.OpenXml.Spreadsheet.InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("#N/A"));
        }
        if (kind == "error-cache") source.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
        sheet.Cell(1, 2).SetFormula("TEXTJOIN(\",\",FALSE,A1,\"tail\")");
        sheet.Cell(2, 2).SetFormula("CONCAT(A1,\"tail\")");
        sheet.Cell(3, 2).SetFormula("CONCATENATE(A1,\"tail\")");
        using var input = new MemoryStream(); document.Save(input); input.Position = 0;
        using var reopened = ExcelDocument.Load(input);
        Assert.True(reopened.Calculate() >= 3);
        var cells = reopened.Sheets[0].WorksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
            .Where(c => c.CellReference?.Value is "B1" or "B2" or "B3").ToArray();
        Assert.Equal(3, cells.Length);
        foreach (var cell in cells) {
            Assert.Equal(kind == "error-cache" ? DocumentFormat.OpenXml.Spreadsheet.CellValues.Error : DocumentFormat.OpenXml.Spreadsheet.CellValues.String, cell.DataType!.Value);
            Assert.Equal(kind == "error-cache" ? "#N/A" : cell.CellReference!.Value == "B1" ? "#N/A,tail" : "#N/Atail", cell.CellValue!.Text);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextJoin_storage_prefixes_nested_calls_and_preserves_literal_and_qualified_text(bool materialized) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(1, 1).SetValue("a");
        sheet.Cell(2, 1).SetValue("b");
        document.AddSheet("O'Brien, TEXTJOIN(").Cell(1, 1).SetValue("quoted");
        if (materialized) sheet.Cell(1, 1).GetValue();
        const string authored = "IF(TRUE,TEXTJOIN(\",\",TRUE,A1:A2,\"TEXTJOIN(\",'O''Brien, TEXTJOIN('!A1),_xlfn.TEXTJOIN(\"\",FALSE,\"x\"))";
        const string stored = "IF(TRUE,_xlfn.TEXTJOIN(\",\",TRUE,A1:A2,\"TEXTJOIN(\",'O''Brien, TEXTJOIN('!A1),_xlfn.TEXTJOIN(\"\",FALSE,\"x\"))";
        sheet.Cell(1, 2).SetFormula(authored);
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
        Assert.Equal("a,b,TEXTJOIN(,quoted", reopened.Sheets[0].Cell(1, 2).GetValue<string>());
    }

    [Theory]
    [InlineData("TEXTJOIN(\",\",FALSE,A1:A2)")]
    [InlineData("TEXTJOIN(\",\",FALSE,#N/A)")]
    [InlineData("TEXTJOIN(#N/A,FALSE,\"text\")")]
    [InlineData("TEXTJOIN(\",\",#N/A,\"text\")")]
    [InlineData("CONCAT(A1:A2)")]
    public void Joining_text_propagates_errors_as_typed_errors(string expression) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(1, 1).SetFormula("IFNA(#N/A,#N/A)");
        sheet.Cell(2, 1).SetValue("tail");
        sheet.Cell(1, 2).SetFormula(expression);
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
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(1, 1).SetValue(new string('x', 20000));
        sheet.Cell(2, 1).SetValue(new string('y', 20000));
        sheet.Cell(1, 2).SetFormula("TEXTJOIN(\",\",FALSE,A1:A2)");
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
        ExcelSheet sheet = document.AddSheet("Report");
        Assert.Throws<ArgumentException>(() => sheet.Cell(1, 1).SetFormula(expression));
    }
}
