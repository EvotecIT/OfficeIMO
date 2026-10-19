using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelXmlTableTests {
    [Fact]
    public void Positional_tables_keep_literal_text_empty_and_missing_cells_without_formula_execution() {
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Data");
        string range = sheet.FromXml("<?xml version='1.0' encoding='UTF-16'?><Table><![CDATA[table label]]><Row><Data>00127</Data><Data>=SUM(A1)</Data><Data></Data></Row><Row><Data>Łódź &amp; &lt;data&gt;</Data></Row></Table>", startRow: 2, startColumn: 3);
        Assert.Equal("C2:E3", range); Assert.Equal("00127", sheet.CellAt(2, 3).GetValue<string>());
        Assert.Equal("=SUM(A1)", sheet.CellAt(2, 4).GetValue<string>()); Assert.Null(sheet.GetFormulaText(2, 4));
        string xml = sheet.ToXml(range);
        using var copy = ExcelDocument.Create(); var target = copy.AddWorksheet("Data");
        Assert.Equal("A1:C2", target.FromXml(xml));
        Assert.Equal("00127", target.CellAt(1, 1).GetValue<string>());
        Assert.Equal("Łódź & <data>", target.CellAt(2, 1).GetValue<string>());
        Assert.Contains("nil=\"true\"", xml);
        using var reopened = ExcelDocument.Load(new MemoryStream(copy.ToBytes()));
        Assert.Equal("=SUM(A1)", reopened.Sheets[0].CellAt(1, 2).GetValue<string>());
    }

    [Fact]
    public void Named_columns_support_namespaces_attributes_explicit_types_and_absent_fields() {
        var options = Mapping();
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Mapped");
        string source = "<orders xmlns='urn:orders'><order id='00127'><total>12.5</total><paid>true</paid><date>2026-10-19T12:30:00</date><ignored>Outside the mapping</ignored></order><order id='00002'><total>0</total><paid>0</paid></order></orders>";
        Assert.Equal("A1:D3", sheet.FromXml(source, includeHeaders: true, options: options));
        Assert.Equal("Identifier", sheet.CellAt(1, 1).GetValue<string>()); Assert.Equal("00127", sheet.CellAt(2, 1).GetValue<string>());
        Assert.Equal(12.5, sheet.CellAt(2, 2).GetValue<double>()); Assert.True(sheet.CellAt(2, 3).GetValue<bool>());
        var read = new ExcelReadOptions { TreatDatesUsingNumberFormat = true };
        string xml = sheet.ToXml("A1:D3", true, options, read);
        XNamespace ns = "urn:orders"; var rows = XElement.Parse(xml).Elements(ns + "order").ToArray();
        Assert.Equal("00127", (string?)rows[0].Attribute("id")); Assert.Equal("true", rows[0].Element(ns + "paid")!.Value);
        Assert.Equal("2026-10-19T12:30:00", rows[0].Element(ns + "date")!.Value); Assert.Null(rows[1].Element(ns + "date"));
        using var copy = ExcelDocument.Create(); var target = copy.AddWorksheet("Mapped"); target.FromXml(xml, options: options);
        Assert.Equal(12.5, target.CellAt(1, 2).GetValue<double>()); Assert.Empty(copy.ValidateOpenXml());
    }

    [Theory]
    [InlineData("<orders xmlns='urn:orders'><order id='1'><total>NaN</total></order></orders>")]
    [InlineData("<orders xmlns='urn:orders'><order><total>1</total><total>2</total></order></orders>")]
    [InlineData("<orders xmlns='urn:orders'><order><total><nested>1</nested></total></order></orders>")]
    [InlineData("<orders xmlns='urn:orders'><wrong/></orders>")]
    public void Invalid_mapping_input_is_rejected_before_replacing_destination_cells(string xml) {
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Data"); sheet.CellValue(1, 1, "Keep");
        Assert.Throws<InvalidDataException>(() => sheet.FromXml(xml, options: Mapping())); Assert.Equal("Keep", sheet.CellAt(1, 1).GetValue<string>());
    }

    [Fact]
    public void Xml_table_resource_limits_cancellation_and_stream_ownership_are_enforced() {
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Data");
        byte[] bytes = Encoding.UTF8.GetBytes("<Table><Row><Data>A</Data><Data>B</Data></Row></Table>");
        using var input = new MemoryStream(bytes); sheet.FromXml(input); Assert.True(input.CanRead);
        string xml = Encoding.UTF8.GetString(bytes);
        Assert.Throws<InvalidDataException>(() => sheet.FromXml(xml, options: new() { MaxInputBytes = 5 }));
        Assert.Throws<InvalidDataException>(() => sheet.FromXml(xml, options: new() { MaxColumns = 1 }));
        Assert.Throws<InvalidDataException>(() => sheet.FromXml(xml, options: new() { MaxCells = 1 }));
        Assert.Throws<InvalidDataException>(() => sheet.FromXml("<Table><Row/><Row/></Table>", options: new() { MaxRows = 1 }));
        Assert.Throws<InvalidDataException>(() => sheet.FromXml(xml, options: new() { MaxXmlDepth = 1 }));
        Assert.Throws<OperationCanceledException>(() => sheet.FromXml(xml, options: new() { CancellationToken = new CancellationToken(true) }));
        Assert.Throws<XmlException>(() => sheet.FromXml("<!DOCTYPE Table [<!ENTITY x SYSTEM 'file:///not-readable'>]><Table><Row><Data>&x;</Data></Row></Table>"));
        Assert.Throws<InvalidDataException>(() => sheet.ToXml("A1:B1", options: new() { MaxOutputBytes = 5 }));
        Assert.Throws<ArgumentException>(() => sheet.ToXml("A1:XFD1048576"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Row_quota_counts_direct_rows_and_not_a_same_named_scalar(bool stream) {
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Data");
        var options = new ExcelXmlTableOptions { MaxRows = 1 }; options.Columns.Add(new() { Name = "Value", XmlName = "Row" });
        const string xml = "<Table><Row><Row>value</Row></Row></Table>";
        if (stream) { using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml)); sheet.FromXml(input, options: options); }
        else sheet.FromXml(xml, options: options);
        Assert.Equal("value", sheet.CellAt(1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Cell_object_reads_materialize_existing_deferred_table_and_json_insertions(bool json) {
        using var document = ExcelDocument.Create(); var sheet = document.AddWorksheet("Data");
        if (json) sheet.FromJson("[{\"Code\":\"00127\"}]", includeHeaders: false);
        else {
            using var table = new System.Data.DataTable(); table.Columns.Add("Code", typeof(string)); table.Rows.Add("00127");
            sheet.InsertDataTable(table, includeHeaders: false);
        }
        Assert.Equal("00127", sheet.CellAt(1, 1).GetFormattedText());
        Assert.Equal("00127", sheet.CellAt(1, 1).GetValue<string>());
    }

    private static ExcelXmlTableOptions Mapping() {
        var options = new ExcelXmlTableOptions { RootElementName = "orders", RowElementName = "order", NamespaceUri = "urn:orders" };
        options.Columns.Add(new() { Name = "Identifier", XmlName = "id", IsAttribute = true });
        options.Columns.Add(new() { Name = "Total", XmlName = "total", ValueType = ExcelXmlTableValueType.Number });
        options.Columns.Add(new() { Name = "Paid", XmlName = "paid", ValueType = ExcelXmlTableValueType.Boolean });
        options.Columns.Add(new() { Name = "Date", XmlName = "date", ValueType = ExcelXmlTableValueType.DateTime });
        return options;
    }
}
