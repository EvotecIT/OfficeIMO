using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelWorksheetGridLoaderTests {
    private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Theory]
    [InlineData("", "xmlns=\"" + MainNamespace + "\"")]
    [InlineData("x:", "xmlns:x=\"" + MainNamespace + "\"")]
    public void First_edit_keeps_grid_markup_through_save(string prefix, string declaration) {
        string grid = "<sheetData><row r=\"1\" spans=\"1:3\" ht=\"24\" customHeight=\"1\" x14ac:dyDescent=\"0.25\">"
            + "<c r=\"A1\"><v>1.50</v></c>"
            + "<c r=\"B1\" t=\"str\"><f>IF(A1&gt;1,&quot;a&amp;b&quot;,&quot;c&quot;)</f><v>a&amp;b</v></c>"
            + "<c r=\"C1\"><f t=\"shared\" ref=\"C1:C2\" si=\"0\">A1*2</f><v>3</v></c></row>"
            + "<row r=\"2\" spans=\"1:3\"><c r=\"C2\"><f t=\"shared\" si=\"0\"/><v>0</v></c></row></sheetData>";
        byte[] package = CreatePackage(Worksheet(prefix, declaration, prefix.Length == 0 ? grid : Prefixed(grid)));

        Worksheet saved = EditAndSave(package);

        Row row = saved.Descendants<Row>().First();
        Assert.Equal("1:3", row.Spans!.InnerText);
        Assert.Equal(24d, row.Height!.Value);
        Assert.True(row.CustomHeight!.Value);
        Assert.Equal(0.25d, row.DyDescent!.Value);
        Cell[] cells = row.Elements<Cell>().ToArray();
        Assert.Equal("1.50", cells[0].CellValue!.Text);
        Assert.Equal(CellValues.String, cells[1].DataType!.Value);
        Assert.Equal("IF(A1>1,\"a&b\",\"c\")", cells[1].CellFormula!.Text);
        Assert.Equal("a&b", cells[1].CellValue!.Text);
        Assert.Equal(CellFormulaValues.Shared, cells[2].CellFormula!.FormulaType!.Value);
        Assert.Equal("C1:C2", cells[2].CellFormula!.Reference!.Value);
        Assert.Equal(0U, cells[2].CellFormula!.SharedIndex!.Value);
        Assert.Empty(saved.Descendants<Row>().ElementAt(1).Elements<Cell>().First().CellFormula!.Text);
    }

    [Theory]
    [InlineData("<sheetData><row r=\"1\"><!-- note --><c r=\"A1\"><v>1</v></c></row></sheetData>")]
    [InlineData("<sheetData><row r=\"1\"><c r=\"A1\" cm=\"1\"><v>1</v></c></row></sheetData>")]
    [InlineData("<sheetData>\n<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>\n</sheetData>")]
    [InlineData("<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>1</t></is></c></row></sheetData>")]
    public void Grid_markup_outside_the_plain_shapes_loads_through_the_sdk(string grid) {
        Worksheet saved = EditAndSave(CreatePackage(Worksheet(string.Empty, "xmlns=\"" + MainNamespace + "\"", grid)));

        Assert.Equal("A1", saved.Descendants<Cell>().First().CellReference!.Value);
    }

    [Theory]
    [InlineData("<sheetData><row r=\"1\"><c r=\"A1\"><v>1</c></row></sheetData>")]
    [InlineData("<sheetData><row r=\"1\" r=\"2\"/></sheetData>")]
    [InlineData("<sheetData><row r=\"1\" x14ac:dyDescent=\"1\"/></sheetData>")]
    public void Malformed_grid_still_fails_as_xml(string grid) {
        string worksheet = "<worksheet xmlns=\"" + MainNamespace + "\">" + grid + "</worksheet>";
        using var document = ExcelDocument.Load(new MemoryStream(CreatePackage(worksheet)));

        Assert.Throws<XmlException>(() => document.Sheets[0].Cell(9, 9).SetValue(1d));
    }

    private static string Worksheet(string prefix, string declaration, string grid) =>
        $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<{prefix}worksheet {declaration} "
        + "xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">"
        + $"{grid}<{prefix}pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></{prefix}worksheet>";

    private static string Prefixed(string grid) => grid.Replace("</", "\u0001").Replace("<", "<x:").Replace("\u0001", "</x:");

    private static Worksheet EditAndSave(byte[] package) {
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Load(new MemoryStream(package))) {
            document.Sheets[0].Cell(9, 9).SetValue(1d);
            document.Save(output);
        }
        output.Position = 0;
        using var saved = SpreadsheetDocument.Open(output, false);
        return (Worksheet)saved.WorkbookPart!.WorksheetParts.Single().Worksheet!.CloneNode(true);
    }

    private static byte[] CreatePackage(string worksheetXml) {
        using var package = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            document.AddSheet("Data").Cell(1, 1).SetValue(1d);
            document.Save(package);
        }
        using (var sdk = SpreadsheetDocument.Open(package, true)) {
            using Stream part = sdk.WorkbookPart!.WorksheetParts.Single().GetStream(FileMode.Create, FileAccess.Write);
            byte[] xml = new UTF8Encoding(false).GetBytes(worksheetXml);
            part.Write(xml, 0, xml.Length);
        }
        return package.ToArray();
    }
}
