using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelAValueAggregateTests {
    [Fact]
    public void Numeric_aggregations_ignore_referenced_text_and_booleans_while_scalar_coercion_remains_available() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Values");
        sheet.CellValue(1, 1, true); sheet.CellValue(2, 1, 0.2d);
        sheet.CellValue(1, 2, "25"); sheet.CellValue(2, 2, 0.2d);
        (string Formula, double Expected)[] cases = {
            ("SUM(A1:B2)", 0.4d), ("AVERAGE(A1:B2)", 0.2d), ("MIN(A1:B2)", 0.2d),
            ("MAX(A1:B2)", 0.2d), ("COUNT(A1:B2)", 2d), ("PRODUCT(A1:B2)", 0.04d),
            ("MEDIAN(A1:B2)", 0.2d), ("MODE.SNGL(A1:B2)", 0.2d),
            ("SUMSQ(A1:B2)", 0.08d), ("LARGE(A1:B2,A1)", 0.2d), ("SMALL(A1:B2,A1)", 0.2d),
            ("SUBTOTAL(9,A1:B2)", 0.4d),
            ("SUMIF(A1:B2,\">=0\")", 0.4d), ("SUMIFS(A1:B2,A1:B2,\">=0\")", 0.4d),
            ("AVERAGEIF(A1:B2,\">=0\")", 0.2d), ("MINIFS(A1:B2,A1:B2,\">=0\")", 0.2d),
            ("ABS(A1)", 1d)
        };
        for (int index = 0; index < cases.Length; index++) sheet.CellFormula(index + 1, 4, cases[index].Formula);
        using var stream = new MemoryStream(); document.Save(stream); stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        Assert.Equal(cases.Length, reopened.Calculate());
        for (int index = 0; index < cases.Length; index++)
            Assert.Equal(cases[index].Expected, reopened.Sheets[0].CellAt(index + 1, 4).GetValue<double>(), 10);
    }

    [Theory]
    [InlineData("AND(TRUE,TRUE)")]
    [InlineData("OR(FALSE,TRUE)")]
    [InlineData("NOT(FALSE)")]
    [InlineData("ISNUMBER(2)")]
    [InlineData("IF(TRUE,TRUE,FALSE)")]
    [InlineData("IFERROR(#N/A,TRUE)")]
    [InlineData("TRUE")]
    public void Boolean_formula_results_keep_their_type_through_references_and_saved_caches(string formula) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Values");
        sheet.CellFormula(1, 1, formula); sheet.CellValue(2, 1, 2d);
        sheet.CellFormula(1, 2, "A1");
        sheet.CellFormula(1, 4, "SUM(A1:A2)");
        sheet.CellFormula(2, 4, "COUNT(A1:A2)");
        sheet.CellFormula(3, 4, "MINA(A1:A2)");
        sheet.CellFormula(4, 4, "SUM(B1,2)");
        sheet.CellFormula(5, 4, "SUM(TRUE,2)");
        Assert.Equal(7, document.Calculate());
        Assert.Equal(2d, sheet.CellAt(1, 4).GetValue<double>());
        Assert.Equal(1d, sheet.CellAt(2, 4).GetValue<double>());
        Assert.Equal(1d, sheet.CellAt(3, 4).GetValue<double>());
        Assert.Equal(2d, sheet.CellAt(4, 4).GetValue<double>());
        Assert.Equal(3d, sheet.CellAt(5, 4).GetValue<double>());
        using var stream = new MemoryStream(); document.Save(stream); stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        var cells = reopened.Sheets[0].WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
        Assert.Equal(CellValues.Boolean, cells["A1"].DataType!.Value);
        Assert.Equal(CellValues.Boolean, cells["B1"].DataType!.Value);
        Assert.Equal("1", cells["A1"].CellValue!.Text);
        Assert.Equal(7, reopened.Calculate());
        Assert.Equal(2d, reopened.Sheets[0].CellAt(1, 4).GetValue<double>());
    }

    [Fact]
    public void Selected_numeric_looking_text_remains_text_and_information_functions_distinguish_booleans() {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Values");
        sheet.CellValue(1, 1, "25");
        sheet.CellFormula(1, 2, "IF(TRUE,A1,0)");
        sheet.CellFormula(1, 3, "B1");
        (string Formula, double Expected)[] cases = {
            ("SUM(C1,2)", 2d), ("MINA(C1,2)", 0d), ("ISTEXT(B1)", 1d),
            ("ISNUMBER(B1)", 0d), ("ISNUMBER(TRUE)", 0d), ("ISNUMBER(AND(TRUE,TRUE))", 0d)
        };
        for (int index = 0; index < cases.Length; index++) sheet.CellFormula(index + 1, 4, cases[index].Formula);
        Assert.Equal(cases.Length + 2, document.Calculate());
        using var stream = new MemoryStream(); document.Save(stream); stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        var cells = reopened.Sheets[0].WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
        Assert.Equal(CellValues.String, cells["B1"].DataType!.Value);
        Assert.Equal(CellValues.String, cells["C1"].DataType!.Value);
        Assert.Equal("25", cells["C1"].CellValue!.Text);
        for (int index = 0; index < cases.Length; index++)
            Assert.Equal(cases[index].Expected, reopened.Sheets[0].CellAt(index + 1, 4).GetValue<double>());
    }

    [Theory]
    [InlineData("MINA", 2d, 1d, 0d)]
    [InlineData("MAXA", 0.2d, 1d, 0.2d)]
    [InlineData("AVERAGEA", 0.2d, 0.6d, 0.1d)]
    public void A_value_aggregates_use_referenced_boolean_and_text_types_and_ignore_blank_cells(string function, double numeric, double logical, double text) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Values");
        sheet.CellValue(1, 1, true); sheet.CellValue(2, 1, numeric);
        sheet.CellValue(1, 2, "25"); sheet.CellValue(2, 2, 0.2d);
        sheet.CellFormula(1, 4, function + "(A1:A3)");
        sheet.CellFormula(2, 4, function + "(B1:B3)");
        sheet.CellFormula(3, 4, "SUM(A1:A3)");
        using var stream = new MemoryStream(); document.Save(stream); stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        Assert.Equal(3, reopened.Calculate());
        Assert.Equal(logical, reopened.Sheets[0].CellAt(1, 4).GetValue<double>(), 10);
        Assert.Equal(text, reopened.Sheets[0].CellAt(2, 4).GetValue<double>(), 10);
        Assert.Equal(numeric, reopened.Sheets[0].CellAt(3, 4).GetValue<double>(), 10);
    }

    [Theory]
    [InlineData("MINA", null)]
    [InlineData("MAXA", null)]
    [InlineData("AVERAGEA", "#DIV/0!")]
    public void A_value_aggregates_resolve_empty_ranges_and_preserve_typed_errors(string function, string? emptyError) {
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Values");
        sheet.CellFormula(1, 1, "IFNA(#N/A,#N/A)");
        sheet.CellFormula(1, 4, function + "(B1:B3)");
        sheet.CellFormula(2, 4, function + "(A1:A3)");
        Assert.Equal(3, document.Calculate());
        using var stream = new MemoryStream(); document.Save(stream); stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        var cells = reopened.Sheets[0].WorksheetPart.Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
        Assert.Equal(emptyError == null ? CellValues.Number : CellValues.Error, cells["D1"].DataType!.Value);
        Assert.Equal(emptyError ?? "0", cells["D1"].CellValue!.Text);
        Assert.Equal(CellValues.Error, cells["D2"].DataType!.Value);
        Assert.Equal("#N/A", cells["D2"].CellValue!.Text);
    }
}
