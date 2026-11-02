using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelWorksheetGridWriterTests {
    private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string DyDescentNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

    [Theory]
    [InlineData("")]
    [InlineData("x")]
    public void Grid_writer_output_is_the_xml_the_sdk_writes(string prefix) {
        string p = prefix.Length == 0 ? string.Empty : prefix + ":";
        string worksheetXml = $"<{p}worksheet xmlns{(prefix.Length == 0 ? string.Empty : ":" + prefix)}=\"{MainNamespace}\" xmlns:x14ac=\"{DyDescentNamespace}\">"
            + $"<{p}sheetData>"
            + $"<{p}row r=\"1\" spans=\"1:4\" s=\"2\" customFormat=\"1\" ht=\"24\" customHeight=\"1\" x14ac:dyDescent=\"0.25\">"
            + $"<{p}c r=\"A1\" s=\"1\"><{p}v>1.50</{p}v></{p}c>"
            + $"<{p}c r=\"B1\" t=\"str\"><{p}f>IF(A1&gt;1,\"a&amp;b\",\"c\")</{p}f><{p}v>a&amp;b &lt;Zażółć — 😀&gt;</{p}v></{p}c>"
            + $"<{p}c r=\"C1\"><{p}f t=\"shared\" ref=\"C1:C2\" ca=\"1\" si=\"0\">A1*2</{p}f><{p}v>3</{p}v></{p}c>"
            + $"<{p}c r=\"D1\" s=\"3\"/></{p}row>"
            + $"<{p}row r=\"2\" hidden=\"1\" outlineLevel=\"1\" collapsed=\"1\"/>"
            + $"<{p}row r=\"3\"><{p}c r=\"A3\" t=\"e\"><{p}v>#N/A</{p}v></{p}c><{p}c r=\"B3\" t=\"b\"><{p}v>1</{p}v></{p}c>"
            + $"<{p}c r=\"C3\"><{p}f t=\"array\" ref=\"C3:C3\">SUM(A1:B1)</{p}f></{p}c></{p}row>"
            + $"</{p}sheetData><{p}pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></{p}worksheet>";
        var worksheet = new Worksheet(worksheetXml);
        // A cell added to a parsed row, as an edit adds one, carries the SDK's own prefix.
        worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Last().AppendChild(new Cell(new CellValue("7")) { CellReference = "D3" });

        var buffer = new ExcelPartBuffer();
        try {
            Assert.True(ExcelWorksheetGridWriter.TrySave(worksheet, buffer));
            string written = Encoding.UTF8.GetString(buffer.Buffer, 0, buffer.ContentLength);

            using var expected = new MemoryStream();
            worksheet.Save(expected);
            AssertSameXml(Encoding.UTF8.GetString(expected.ToArray()), written);
        } finally {
            buffer.Release();
        }
    }

    [Theory]
    [InlineData("<c r=\"A1\" t=\"inlineStr\"><is><t>text</t></is></c>")]
    [InlineData("<c r=\"A1\" cm=\"1\"><v>1</v></c>")]
    [InlineData("<c r=\"A1\" t=\"str\"><v>tab\there</v></c>")]
    [InlineData("<c r=\"A1\"><f t=\"shared\" si=\"0\"/><v>1</v></c>")]
    public void Grid_writer_declines_cells_outside_its_shapes(string cell) {
        var worksheet = new Worksheet($"<worksheet xmlns=\"{MainNamespace}\"><sheetData><row r=\"1\">{cell}</row></sheetData></worksheet>");
        var buffer = new ExcelPartBuffer();
        try {
            Assert.False(ExcelWorksheetGridWriter.TrySave(worksheet, buffer));
        } finally {
            buffer.Release();
        }
    }

    /// <summary>The writer names each attribute it writes; one the SDK gains later would be dropped unless it is listed here.</summary>
    [Theory]
    [InlineData(typeof(Row), "RowIndex,Spans,StyleIndex,CustomFormat,Height,Hidden,CustomHeight,OutlineLevel,Collapsed,ThickTop,ThickBot,ShowPhonetic,DyDescent")]
    [InlineData(typeof(Cell), "CellReference,StyleIndex,DataType,CellMetaIndex,ValueMetaIndex,ShowPhonetic")]
    [InlineData(typeof(CellFormula), "FormulaType,AlwaysCalculateArray,Reference,DataTable2D,DataTableRow,Input1Deleted,Input2Deleted,R1,R2,CalculateCell,SharedIndex,Bx,Space")]
    [InlineData(typeof(CellValue), "Space")]
    public void Grid_writer_knows_every_attribute_of_the_grid_elements(Type type, string handled) {
        string[] attributes = type.GetProperties()
            .Where(static property => typeof(OpenXmlSimpleType).IsAssignableFrom(property.PropertyType))
            .Select(static property => property.Name)
            .OrderBy(static name => name, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(handled.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).OrderBy(static name => name, StringComparer.Ordinal), attributes);
    }

    /// <summary>Compares element names, attributes and text, ignoring prefixes and namespace declarations.</summary>
    private static void AssertSameXml(string expected, string actual) {
        Assert.Equal(Describe(XElement.Parse(expected)), Describe(XElement.Parse(actual)));
    }

    private static string Describe(XElement element) {
        var builder = new StringBuilder();
        Describe(element, builder);
        return builder.ToString();
    }

    private static void Describe(XElement element, StringBuilder builder) {
        builder.Append('<').Append(element.Name);
        foreach (XAttribute attribute in element.Attributes().Where(static attribute => !attribute.IsNamespaceDeclaration)) {
            builder.Append(' ').Append(attribute.Name).Append("=\"").Append(attribute.Value).Append('"');
        }
        builder.Append('>');
        if (!element.HasElements) builder.Append(element.Value);
        foreach (XElement child in element.Elements()) Describe(child, builder);
        builder.Append("</>");
    }
}
