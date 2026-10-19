using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelSpreadsheetXmlTests {
    private const string Ss = "urn:schemas-microsoft-com:office:spreadsheet";
    private static byte[] Workbook(string content) => Encoding.UTF8.GetBytes("<Workbook xmlns=\"" + Ss + "\" xmlns:ss=\"" + Ss + "\">" + content + "</Workbook>");
    private static byte[] Cells(string content) => Workbook("<Worksheet ss:Name=\"Data\"><Table><Row>" + content + "</Row></Table></Worksheet>");

    [Fact]
    public void Xml_import_reads_native_Excel_saved_formulas_and_caches_without_calculation() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "SpreadsheetXmlCorpus", "excel-financial.xml");
        using var document = ExcelDocument.ImportSpreadsheetXml(path).Value;
        Assert.Equal(38, document.InspectFormulas().Formulas.Count);
        Assert.InRange(document.Sheets[0].CellAt(1, 1).GetValue<double>(), 0.00770147248, 0.00770147250);
        Assert.Equal("#NUM!", document.InspectFormulas().Formulas.Single(formula => formula.CellReference == "A14").CachedValue);
        using var copy = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal("RATE(48,-200,8000)", copy.Sheets[0].GetFormulaText(1, 1));
        Assert.Empty(copy.ValidateOpenXml());
    }

    [Fact]
    public void Xml_import_retains_types_sparse_positions_formula_caches_and_editable_workbook() {
        byte[] source = Workbook("""
            <Styles><Style ss:ID="date"><NumberFormat ss:Format="yyyy-mm-dd"/></Style></Styles>
            <Worksheet ss:Name="Data"><Table><Row ss:Index="3">
              <Cell ss:Index="2"><Data ss:Type="String">  00127 &amp; Łódź  </Data></Cell>
              <Cell><Data ss:Type="Number">1.2345678901234567</Data></Cell>
              <Cell><Data ss:Type="Boolean">1</Data></Cell>
              <Cell ss:StyleID="date"><Data ss:Type="DateTime">2026-10-19T12:30:00.000</Data></Cell>
              <Cell ss:Formula="=RC[-3]*2"><Data ss:Type="Number">99</Data></Cell>
              <Cell><Data ss:Type="Error">#N/A</Data></Cell>
              <Cell ss:Formula="=&quot;RC[-3]&quot;"><Data ss:Type="String">RC[-3]</Data></Cell>
            </Row></Table></Worksheet><Worksheet ss:Name="Empty"><Table/></Worksheet>
            """);
        var result = ExcelDocument.ImportSpreadsheetXml(source);
        using var document = result.RequireNoLoss();
        Assert.Equal(2, document.Sheets.Count);
        var sheet = document.Sheets[0];
        Assert.Equal("  00127 & Łódź  ", sheet.CellAt(3, 2).GetValue<string>());
        Assert.Equal(1.2345678901234567, sheet.CellAt(3, 3).GetValue<double>());
        Assert.True(sheet.CellAt(3, 4).GetValue<bool>());
        Assert.Equal("C3*2", sheet.GetFormulaText(3, 6));
        Assert.Equal(99, sheet.CellAt(3, 6).GetValue<double>());
        Assert.Equal("\"RC[-3]\"", sheet.GetFormulaText(3, 8));
        Assert.Equal("#N/A", document.OpenXmlDocument.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single(cell => cell.CellReference == "G3").CellValue!.Text);
        sheet.CellValue(4, 1, "Edited");
        using var saved = new MemoryStream(); document.Save(saved);
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal("Edited", reopened.Sheets[0].CellAt(4, 1).GetValue<string>());
        using var reader = ExcelDocument.OpenDataReader(saved.ToArray(), new ExcelReadOptions { SheetName = "Data", A1Range = "E3:E3", HasHeaderRow = false, TreatDatesUsingNumberFormat = true });
        Assert.True(reader.Read());
        Assert.Equal(new DateTime(2026, 10, 19, 12, 30, 0), reader.GetDateTime(0));
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Fact]
    public void Xml_export_reimports_cells_and_formula_references_without_evaluating_them() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("A & B");
        sheet.CellValue(2, 2, "00127"); sheet.CellValue(2, 3, false); sheet.CellValue(2, 4, 12.5);
        sheet.CellValue(4, 4, 42d); sheet.CellFormula(4, 4, "D2*2");
        var result = document.ExportSpreadsheetXml();
        Assert.Equal("Workbook", XDocument.Parse(Encoding.UTF8.GetString(result.Value)).Root!.Name.LocalName);
        var imported = ExcelDocument.ImportSpreadsheetXml(result.Value);
        using var copy = imported.RequireValue();
        Assert.Equal("A & B", copy.Sheets[0].Name);
        Assert.Equal("00127", copy.Sheets[0].CellAt(2, 2).GetValue<string>());
        Assert.False(copy.Sheets[0].CellAt(2, 3).GetValue<bool>());
        Assert.Equal(12.5, copy.Sheets[0].CellAt(2, 4).GetValue<double>());
        Assert.Equal("D2*2", copy.Sheets[0].GetFormulaText(4, 4));
        Assert.Equal(42d, copy.Sheets[0].CellAt(4, 4).GetValue<double>());
        Assert.Empty(copy.ValidateOpenXml());
    }

    [Fact]
    public void Xml_text_keeps_carriage_returns_through_export_import_and_package_save() {
        const string text = "A\rB\r\nC\nD";
        using var document = ExcelDocument.Create();
        document.AddWorksheet("Data").CellValue(1, 1, text);
        byte[] xml = document.ExportSpreadsheetXml().Value;
        using var imported = ExcelDocument.ImportSpreadsheetXml(xml).Value;
        Assert.Equal(text, imported.Sheets[0].CellAt(1, 1).GetValue<string>());
        using var reopened = ExcelDocument.Load(new MemoryStream(imported.ToBytes()));
        Assert.Equal(text, reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Xml_export_preserves_caches_and_never_invokes_save_calculation(bool evaluate) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, 99d); sheet.CellFormula(1, 1, "EXPORTCALL()");
        int calls = 0;
        document.Calculation.RegisterCustomFunction("EXPORTCALL", (_, _) => { calls++; return ExcelFormulaValue.FromNumber(5d); });
        document.Calculation.EvaluateFormulasBeforeSave = evaluate;
        document.Calculation.ClearCachedFormulaResultsBeforeSave = !evaluate;
        using var imported = ExcelDocument.ImportSpreadsheetXml(document.ExportSpreadsheetXml().Value).Value;
        Assert.Equal(0, calls);
        Assert.Equal(99d, sheet.CellAt(1, 1).GetValue<double>());
        Assert.Equal(99d, imported.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(evaluate, document.Calculation.EvaluateFormulasBeforeSave);
        Assert.Equal(!evaluate, document.Calculation.ClearCachedFormulaResultsBeforeSave);
        // These policies still apply to the ordinary save route.
        _ = document.ToBytes();
        Assert.Equal(evaluate, calls > 0);
        if (evaluate) Assert.Equal(5d, sheet.CellAt(1, 1).GetValue<double>());
        else Assert.Null(document.OpenXmlDocument.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single().CellValue);
    }

    [Fact]
    public void Xml_export_accepts_read_only_workbooks_and_does_not_require_saving_the_source() {
        using var source = ExcelDocument.Create(); source.AddWorksheet("Data").CellValue(1, 1, "Read only");
        using var document = ExcelDocument.Load(new MemoryStream(source.ToBytes()), new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        byte[] output = document.ExportSpreadsheetXml().Value;
        using var imported = ExcelDocument.ImportSpreadsheetXml(output).Value;
        Assert.Equal("Read only", imported.Sheets[0].CellAt(1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(false, "1900-01-01T06:00:00", 1.25)]
    [InlineData(false, "1900-03-01T06:00:00", 61.25)]
    [InlineData(true, "1904-01-01T06:00:00", 0.25)]
    public void Xml_dates_use_the_shared_excel_epoch_contract(bool date1904, string date, double expected) {
        byte[] input = Workbook((date1904 ? "<x:ExcelWorkbook xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><x:Date1904/></x:ExcelWorkbook>" : "")
            + "<Worksheet ss:Name=\"Date\"><Table><Row><Cell><Data ss:Type=\"DateTime\">" + date + "</Data></Cell></Row></Table></Worksheet>");
        using var document = ExcelDocument.ImportSpreadsheetXml(input).Value;
        Assert.Equal(expected, document.Sheets[0].CellAt(1, 1).GetValue<double>());
        using var copy = ExcelDocument.ImportSpreadsheetXml(document.ExportSpreadsheetXml().Value).Value;
        Assert.Equal(expected, copy.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Fact]
    public void Omitted_features_are_reported_and_strict_consumers_can_reject_the_result() {
        var result = ExcelDocument.ImportSpreadsheetXml(Workbook("""
            <Styles><Style ss:ID="Default"><Font ss:Bold="1"/></Style></Styles>
            <Worksheet ss:Name="Data"><Table><Row><Cell ss:MergeAcross="1"><Data ss:Type="String">A</Data></Cell>
            <Cell><Data ss:Type="String">C</Data></Cell></Row></Table></Worksheet>
            """));
        using var document = result.Value;
        Assert.True(result.HasLoss);
        Assert.Throws<OfficeConversionException>(() => result.RequireNoLoss());
        Assert.Equal("C", document.Sheets[0].CellAt(1, 3).GetValue<string>());
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Message.Contains("Font"));
        Assert.Contains(result.Report.FidelityDiagnostics, item => item.Message.Contains("merged-cell"));
        document.Sheets[0].AddTable("A1:C1", hasHeader: true, name: "DataTable", style: ExcelTableStyle.TableStyleMedium9);
        Assert.True(document.ExportSpreadsheetXml().HasLoss);
    }

    [Theory]
    [InlineData("<Cell><Data ss:Type=\"Number\">NaN</Data></Cell>")]
    [InlineData("<Cell><Data ss:Type=\"Boolean\">true</Data></Cell>")]
    [InlineData("<Cell><Data ss:Type=\"Error\">bogus</Data></Cell>")]
    [InlineData("<Cell ss:Index=\"257\"/>")]
    [InlineData("<Cell ss:Index=\"2\"/><Cell ss:Index=\"1\"/>")]
    [InlineData("<Cell ss:StyleID=\"missing\"/>")]
    public void Invalid_types_coordinates_or_styles_fail_instead_of_changing_values(string cells) =>
        Assert.Throws<InvalidDataException>(() => ExcelDocument.ImportSpreadsheetXml(Cells(cells)));

    [Fact]
    public void Limits_cancellation_and_caller_stream_ownership_apply_before_returning_a_workbook() {
        byte[] source = Cells("<Cell><Data ss:Type=\"String\">Hello</Data></Cell>");
        using var stream = new MemoryStream(source);
        using var document = ExcelDocument.ImportSpreadsheetXml(stream).Value;
        Assert.True(stream.CanRead); Assert.Equal(stream.Length, stream.Position);
        Assert.Throws<InvalidDataException>(() => ExcelDocument.ImportSpreadsheetXml(source, new() { MaxInputBytes = 10 }));
        Assert.Throws<InvalidDataException>(() => ExcelDocument.ImportSpreadsheetXml(Cells("<Cell/><Cell/>"), new() { MaxCells = 1 }));
        Assert.Throws<InvalidDataException>(() => ExcelDocument.ImportSpreadsheetXml(source, new() { MaxXmlDepth = 2 }));
        Assert.Throws<InvalidDataException>(() => ExcelDocument.ImportSpreadsheetXml(source, new() { MaxXmlElements = 2 }));
        Assert.Throws<OperationCanceledException>(() => ExcelDocument.ImportSpreadsheetXml(source, new() { CancellationToken = new CancellationToken(true) }));
        Assert.Throws<InvalidDataException>(() => document.ExportSpreadsheetXml(new() { MaxOutputBytes = 16 }));
        Assert.Throws<XmlException>(() => ExcelDocument.ImportSpreadsheetXml(Encoding.UTF8.GetBytes("<!DOCTYPE Workbook [<!ENTITY x SYSTEM 'file:///does-not-exist'>]><Workbook xmlns='" + Ss + "'>&x;</Workbook>")));
    }

    [Theory]
    [InlineData("SUM(R[-2]C:R[-1]C)", "SUM(D2:D3)")]
    [InlineData("'Other sheet'!R1C1+RC[-1]", "'Other sheet'!$A$1+C4")]
    [InlineData("IF(RC[-1]=1,\"R1C1\",\"a\"\"RC\")", "IF(C4=1,\"R1C1\",\"a\"\"RC\")")]
    [InlineData("COUNT(R1C1:R2C2)+RATE(1,2,3)", "COUNT($A$1:$B$2)+RATE(1,2,3)")]
    [InlineData("'R1C1'!LoanRate+RC[-1]", "'R1C1'!LoanRate+C4")]
    [InlineData("'RC taxes'!LoanRate+'Bob''s RC plan'!LoanRate", "'RC taxes'!LoanRate+'Bob''s RC plan'!LoanRate")]
    [InlineData("'[R1C1.xls]RC plan'!LoanRate+'RC taxes'!RC[-1]", "'[R1C1.xls]RC plan'!LoanRate+'RC taxes'!C4")]
    public void R1C1_formula_conversion_preserves_literals_functions_and_qualified_references(string formula, string expected) {
        var syntax = ExcelFormulaSyntaxTree.Parse(formula, ExcelReferenceStyle.R1C1, 4, 4);
        Assert.Equal(formula, syntax.Text);
        Assert.Equal(expected, syntax.ConvertReferences(ExcelReferenceStyle.A1));
    }
}
