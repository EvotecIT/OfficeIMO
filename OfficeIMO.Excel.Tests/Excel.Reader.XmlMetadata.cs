using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
#if NET8_0_OR_GREATER
    [Fact]
    public void StylesValidationObservesCancellationAfterCellFormats() {
        string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs><cellStyles>"
            + string.Concat(Enumerable.Repeat("<cellStyle name=\"Default\" xfId=\"0\"/>", 1000))
            + "</cellStyles></styleSheet>";
        using var cancellation = new System.Threading.CancellationTokenSource();
        using var input = new TailCancellationInput(Encoding.UTF8.GetBytes(styles), 4096, cancellation) { Armed = true };
        var provider = new StylesCacheProvider(() => input);

        var exception = Assert.Throws<OperationCanceledException>(() => provider.GetValue(cancellation.Token));
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(0, input.ReadsAfterCancellation);
    }
#endif

    [Theory]
    [InlineData(false, "trailing-markup")]
    [InlineData(true, "trailing-markup")]
    [InlineData(false, "entity")]
    [InlineData(true, "entity")]
    [InlineData(false, "character")]
    [InlineData(true, "character")]
    [InlineData(false, "duplicate-attribute")]
    [InlineData(true, "duplicate-attribute")]
    public void OpenDataReader_RejectsMalformedStylesAfterCellFormats(bool streamedPart, string damage) {
        string path = CreateCompactFastPathWorkbook();
        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes("""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
                  <row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c></row>
                  <row r="2"><c r="A2" s="0"><v>42</v></c></row>
                  <row r="3"><c r="A3" s="0"><v>43</v></c></row>
                </sheetData></worksheet>
                """));
            string padding = streamedPart ? new string('x', 10_000) : "Default";
            string styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs>"
                + "<cellStyles count=\"1\"><cellStyle name=\"" + padding
                + "\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(styles));
            AssertCompactNumericRows(path);

            string malformed = damage switch {
                "trailing-markup" => styles + "<",
                "entity" => styles.Replace("name=\"" + padding + "\"", "name=\"" + padding + "&missing;\""),
                "character" => styles.Replace("name=\"" + padding + "\"", "name=\"" + padding + "bad\u0001text\""),
                "duplicate-attribute" => styles.Replace("name=\"" + padding + "\"", "name=\"" + padding + "\" name=\"second\""),
                _ => throw new ArgumentOutOfRangeException(nameof(damage))
            };
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(malformed));
            Assert.Throws<System.Xml.XmlException>(() => {
                using var reader = ExcelDocument.OpenReader(path);
                _ = reader.FieldCount;
            });
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("plain", false)]
    [InlineData("plain", true)]
    [InlineData("entities", false)]
    [InlineData("entities", true)]
    [InlineData("padded-style", false)]
    [InlineData("padded-style", true)]
    [InlineData("long-style", false)]
    [InlineData("long-style", true)]
    [InlineData("unknown-type", false)]
    [InlineData("unknown-type", true)]
    [InlineData("alternate-formats", false)]
    [InlineData("alternate-formats", true)]
    [InlineData("process-number-formats", false)]
    [InlineData("process-number-formats", true)]
    [InlineData("process-cell-formats", false)]
    [InlineData("process-cell-formats", true)]
    [InlineData("process-root", false)]
    [InlineData("process-root", true)]
    [InlineData("process-local-scope", false)]
    [InlineData("process-local-scope", true)]
    [InlineData("process-prefix-scope", false)]
    [InlineData("process-prefix-scope", true)]
    [InlineData("process-wildcard", false)]
    [InlineData("process-wildcard", true)]
    public void Reader_XmlMetadataPreservesTypedAndMaterializedValues(string shape, bool date1904) {
        string path = CreateCompactFastPathWorkbook();
        try {
            using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, true)) {
                workbook.WorkbookPart!.Workbook.WorkbookProperties =
                    new DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties { Date1904 = date1904 };
            }
            string numberType = shape == "entities" ? "&#110;" : shape == "unknown-type" ? new string('x', 31) + "😀tail" : "n";
            string booleanType = shape == "entities" ? "&#98;" : "b";
            string style = shape switch {
                "entities" => "&#49;",
                "padded-style" => " +1 ",
                "long-style" => new string('0', 80) + "1",
                _ => "1"
            };
            string xml = $$"""
                <x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData>
                  <x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Id</x:t></x:is></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>Active</x:t></x:is></x:c><x:c r="C1" t="inlineStr"><x:is><x:t>Name</x:t></x:is></x:c><x:c r="D1" t="inlineStr"><x:is><x:t>Created</x:t></x:is></x:c></x:row>
                  <x:row r="2"><x:c r="A2" t="{{numberType}}" s="0"><x:v>42</x:v></x:c><x:c r="B2" t="{{booleanType}}"><x:v>1</x:v></x:c><x:c r="C2" t="str"><x:v>Alpha</x:v></x:c><x:c r="D2" t="n" s="{{style}}"><x:v>45351</x:v></x:c></x:row>
                  <x:row r="3"><x:c r="A3" t="n" s="0"><x:v>43</x:v></x:c><x:c r="B3" t="b"><x:v>0</x:v></x:c><x:c r="C3" t="str"><x:v>Beta</x:v></x:c><x:c r="D3" s="1"><x:v>45352</x:v></x:c></x:row>
                </x:sheetData></x:worksheet>
                """;
            ReplaceXmlMetadataWorksheet(path, xml);
            string styles = shape == "alternate-formats" ? """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p="urn:private:future-style">
                  <numFmts count="1"><mc:AlternateContent>
                    <mc:Choice Requires="p"><numFmt numFmtId="165" formatCode="0.00"/></mc:Choice>
                    <mc:Fallback><numFmt numFmtId="165" formatCode="yyyy-mm-dd"/></mc:Fallback>
                  </mc:AlternateContent></numFmts>
                  <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="165" applyNumberFormat="1"/></cellXfs>
                </styleSheet>
                """ : """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>
                """;
            if (shape.StartsWith("process-", StringComparison.Ordinal)) {
                string formats = "<numFmts count=\"1\"><numFmt numFmtId=\"165\" formatCode=\"yyyy-mm-dd\"/></numFmts>";
                string cellFormats = "<cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"165\"/></cellXfs>";
                string processContent = "p:payload";
                switch (shape) {
                    case "process-number-formats":
                        formats = formats.Replace("<numFmt numFmtId", "<p:payload><numFmt numFmtId").Replace("</numFmts>", "</p:payload></numFmts>");
                        break;
                    case "process-cell-formats":
                        cellFormats = cellFormats.Replace("<xf numFmtId=\"165\"/>", "<p:payload><xf numFmtId=\"165\"/></p:payload>");
                        break;
                    case "process-root":
                        formats = "<p:payload>" + formats;
                        cellFormats += "</p:payload>";
                        break;
                    case "process-local-scope":
                        processContent = "";
                        formats = formats.Replace("<numFmts count=\"1\">", "<numFmts count=\"1\" mc:ProcessContent=\"p:payload\"><p:payload>")
                            .Replace("</numFmts>", "</p:payload></numFmts>");
                        cellFormats = cellFormats.Replace("<xf numFmtId=\"165\"/>", "<p:payload><xf numFmtId=\"0\"/></p:payload><xf numFmtId=\"165\"/>");
                        break;
                    case "process-prefix-scope":
                        formats = formats.Replace("<numFmt numFmtId", "<p:payload><numFmt numFmtId").Replace("</numFmts>", "</p:payload></numFmts>");
                        cellFormats = cellFormats.Replace("<cellXfs count=\"2\">", "<cellXfs count=\"2\" xmlns:p=\"urn:private:other-style\" mc:Ignorable=\"p\">")
                            .Replace("<xf numFmtId=\"165\"/>", "<p:payload><xf numFmtId=\"0\"/></p:payload><xf numFmtId=\"165\"/>");
                        break;
                    case "process-wildcard":
                        processContent = "p:*";
                        formats = "<p:numberFormats>" + formats + "</p:numberFormats>";
                        cellFormats = "<p:cellFormats>" + cellFormats + "</p:cellFormats>";
                        break;
                }
                styles = "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                    + " xmlns:p=\"urn:private:future-style\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""
                    + " mc:Ignorable=\"p\" mc:ProcessContent=\"" + processContent + "\">" + formats + cellFormats + "</styleSheet>";
            }
            ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(styles));
            var options = new ExcelReaderOptions { NumericAsDecimal = true, InferDataTableColumnTypes = false };
            object[][] expected = [
                [42m, true, "Alpha", new DateTime(2024, 2, 29).AddDays(date1904 ? 1462 : 0)],
                [43m, false, "Beta", new DateTime(2024, 3, 1).AddDays(date1904 ? 1462 : 0)]
            ];
            using (var reader = ExcelDocument.OpenReader(path, options)) {
                for (int row = 0; row < 2; row++) {
                    Assert.True(reader.Read());
                    Assert.Equal(expected[row][0], reader.GetDecimal(0));
                    Assert.Equal(expected[row][1], reader.GetBoolean(1));
                    Assert.Equal(expected[row][2], reader.GetString(2));
                    Assert.Equal(expected[row][3], reader.GetDateTime(3));
                    for (int column = 0; column < 4; column++) Assert.Equal(expected[row][column], reader.GetValue(column));
                }
                Assert.False(reader.Read());
            }
            using var owner = ExcelDocumentReader.Open(path, options);
            var sheet = owner.GetSheet("Data");
            var values = sheet.ReadRange("A2:D3", ExcelExecutionMode.Sequential);
            using var table = sheet.ReadRangeAsDataTable("A1:D3", headersInFirstRow: true);
            var objects = sheet.ReadObjects<XmlMetadataRecord>("A1:D3").ToArray();
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal(2, objects.Length);
            for (int row = 0; row < 2; row++) {
                for (int column = 0; column < 4; column++) {
                    Assert.Equal(expected[row][column], values[row, column]);
                    Assert.Equal(expected[row][column], table.Rows[row][column]);
                }
                Assert.Equal(expected[row][0], objects[row].Id);
                Assert.Equal(expected[row][1], objects[row].Active);
                Assert.Equal(expected[row][2], objects[row].Name);
                Assert.Equal(expected[row][3], objects[row].Created);
            }
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("")]
    [InlineData("4294967296")]
    [InlineData("-1")]
    [InlineData("long")]
    [InlineData("surrogate")]
    public void DataReader_XmlMetadataRejectsInvalidStylesWithOriginalCellReference(string style) {
        string path = CreateCompactFastPathWorkbook();
        try {
            if (style == "long") style = new string('0', 80) + "x";
            if (style == "surrogate") style = new string('0', 31) + "😀0";
            ReplaceXmlMetadataWorksheet(path, "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + $"<row r=\"2\"><c r=\"&#67;2\" t=\"n\" s=\"{style}\"><v>42</v></c></row></sheetData></worksheet>");
            var error = Assert.Throws<InvalidDataException>(() => {
                using var reader = ExcelDocument.OpenReader(path);
 _ = reader.FieldCount;             });
            Assert.Contains("cell C2 contains an invalid cell style index", error.Message);
        } finally { File.Delete(path); }
    }

    private static void ReplaceXmlMetadataWorksheet(string path, string xml) =>
        ReplaceZipEntry(path, "xl/worksheets/sheet1.xml",
            Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray());

    private sealed class XmlMetadataRecord {
        public decimal Id { get; set; }
        public bool Active { get; set; }
        public string? Name { get; set; }
        public DateTime Created { get; set; }
    }
}
