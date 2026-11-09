using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("numeric", false, false)]
    [InlineData("text-colon", false, false)]
    [InlineData("prefixed-attribute", false, false)]
    [InlineData("strict", false, false)]
    [InlineData("malformed-metadata", true, false)]
    [InlineData("malformed-tail", true, false)]
    [InlineData("duplicate-row-attribute", true, false)]
    [InlineData("unbound-prefix", true, false)]
    [InlineData("normalized-prefix-alias", true, false)]
    [InlineData("malformed-body", true, false)]
    [InlineData("normalized-prefix-alias-canonical", true, false)]
    [InlineData("namespace-alias-canonical-valid", false, false)]
    [InlineData("canonical-invalid-xml-space", true, false)]
    [InlineData("canonical-xml-space", false, false)]
    [InlineData("numeric", false, true)]
    [InlineData("text-colon", false, true)]
    [InlineData("strict", false, true)]
    [InlineData("malformed-tail", true, true)]
    [InlineData("duplicate-row-attribute", true, true)]
    [InlineData("normalized-prefix-alias-canonical", true, true)]
    [InlineData("canonical-invalid-xml-space", true, true)]
    [InlineData("malformed-body", true, true)]
    [InlineData("unbound-prefix", true, true)]
    [InlineData("rebound-value", true, true)]
    [InlineData("mixed-cell-qualification", true, true)]
    [InlineData("unqualified-empty-cell", true, true)]
    [InlineData("mixed-valid-cell", false, true)]
    public void OpenDataReader_PreservesIndexedBodyAndEnclosingXmlValidation(string scenario, bool invalid, bool prefixed) {
        const string spreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        string xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<worksheet xmlns=\"" + spreadsheetNamespace + "\" xmlns:a=\"urn:one\" xmlns:b=\"urn:two\">"
            + "<view:sheetViews xmlns:view=\"" + spreadsheetNamespace + "\"><view:sheetView workbookViewId=\"0\"/></view:sheetViews>"
            + "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Value</t></is></c></row>"
            + "<row r=\"2\" spans=\"1:1\"><c r=\"A2\"><v>42</v></c></row>"
            + "<row r=\"3\"><c r=\"A3\"><v>43</v></c></row></sheetData></worksheet>";
        string original = xml;
        switch (scenario) {
            case "text-colon":
                xml = xml.Replace("<c r=\"A2\"><v>42</v></c>", "<c r=\"A2\" t=\"inlineStr\"><is><t>12:34</t></is></c>");
                break;
            case "prefixed-attribute":
                xml = xml.Replace("spans=\"1:1\"", "spans=\"1:1\" a:note=\"kept\"");
                break;
            case "strict":
                xml = xml.Replace(spreadsheetNamespace, "http://purl.oclc.org/ooxml/spreadsheetml/main");
                break;
            case "malformed-metadata":
                xml = xml.Replace("</view:sheetViews>", "</view:sheetView>");
                break;
            case "malformed-tail":
                xml = xml.Replace("<row ", "\n<row ").Replace("</worksheet>", "<!--invalid--comment--></worksheet>");
                break;
            case "duplicate-row-attribute":
                xml = xml.Replace("r=\"2\" spans", "r=\"2\" r=\"2\" spans");
                break;
            case "unbound-prefix":
                xml = xml.Replace("spans=\"1:1\"", "spans=\"1:1\" missing:note=\"bad\"");
                break;
            case "normalized-prefix-alias":
            case "normalized-prefix-alias-canonical":
                xml = xml.Replace("urn:one", "urn:same\t").Replace("urn:two", "urn:same ")
                    .Replace("spans=\"1:1\"", "spans=\"1:1\" a:note=\"one\" b:note=\"two\"");
                break;
            case "namespace-alias-canonical-valid":
                xml = xml.Replace("spans=\"1:1\"", "spans=\"1:1\" a:note=\"one\" b:note=\"two\"");
                break;
            case "canonical-invalid-xml-space":
                xml = xml.Replace("spans=\"1:1\"", "spans=\"1:1\" xml:space=\"bogus\"");
                break;
            case "canonical-xml-space":
                xml = xml.Replace("spans=\"1:1\"", "spans=\"1:1\" xml:space=\"preserve\"");
                break;
            case "malformed-body":
                xml = xml.Replace("<v>42</v>", "<v>42</missing>");
                break;
            case "mixed-cell-qualification":
            case "unqualified-empty-cell":
            case "mixed-valid-cell":
                xml = xml.Replace(" spans=\"1:1\"", "");
                break;
            case "rebound-value":
                xml = xml.Replace("<v>42</v>", "<v xmlns:q=\"urn:unexpected\">42</v>");
                break;
        }
        if (scenario == "normalized-prefix-alias-canonical"
            || scenario == "namespace-alias-canonical-valid"
            || scenario == "canonical-invalid-xml-space"
            || scenario == "canonical-xml-space") {
            xml = xml.Replace(
                "<view:sheetViews xmlns:view=\"" + spreadsheetNamespace + "\"><view:sheetView workbookViewId=\"0\"/></view:sheetViews>",
                "<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>");
        }
        if (scenario != "numeric") Assert.NotEqual(original, xml);
        if (prefixed) {
            xml = System.Text.RegularExpressions.Regex.Replace(xml,
                @"(</?)(worksheet|sheetData|row|c|v|is|t|sheetViews|sheetView)(?=[\s/>])", "$1q:$2");
            string prefixNamespace = scenario == "strict"
                ? "http://purl.oclc.org/ooxml/spreadsheetml/main" : spreadsheetNamespace;
            xml = xml.Replace("<q:worksheet ", "<q:worksheet xmlns:q=\"" + prefixNamespace + "\" ");
            if (scenario == "mixed-cell-qualification") {
                xml = xml.Replace("<q:c r=\"A2\">", "<c r=\"A2\">");
            } else if (scenario == "unqualified-empty-cell") {
                xml = xml.Replace(" xmlns=\"" + spreadsheetNamespace + "\"", "")
                    .Replace("<q:c r=\"A2\"><q:v>42</q:v></q:c>", "<c r=\"A2\"/>");
            } else if (scenario == "mixed-valid-cell") {
                xml = xml.Replace("<q:c r=\"A2\"><q:v>42</q:v></q:c>", "<c r=\"A2\"><q:v>42</q:v></c>");
            }
        }
        string path = CreateCompactFastPathWorkbook();
        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            if (scenario == "rebound-value" || scenario == "unqualified-empty-cell") {
                InvalidDataException error = Assert.Throws<InvalidDataException>(() => {
                    using var reader = ExcelDocument.OpenReader(path);
                    _ = reader.FieldCount;
                });
                string element = scenario == "rebound-value" ? "v" : "c";
                Assert.Contains("element '" + element + "' must use a SpreadsheetML namespace", error.Message);
                return;
            }
            if (invalid) {
                XmlException error = Assert.Throws<XmlException>(() => {
                    using var reader = ExcelDocument.OpenReader(path);
                    _ = reader.FieldCount;
                });
                if (scenario == "malformed-tail") {
                    XmlException originalError = Assert.Throws<XmlException>(() => {
                        using var originalReader = XmlReader.Create(new StringReader(xml),
                            new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
                        while (originalReader.Read()) { }
                    });
                    Assert.Equal(originalError.LineNumber, error.LineNumber);
                    Assert.Equal(originalError.LinePosition, error.LinePosition);
                }
                if (scenario == "normalized-prefix-alias" || scenario == "normalized-prefix-alias-canonical") {
                    Assert.Equal(1, error.LineNumber);
                    Assert.True(error.LinePosition > xml.IndexOf(prefixed ? "<q:row r=\"2\"" : "<row r=\"2\"", StringComparison.Ordinal));
                }
                return;
            }
            using var data = ExcelDocument.OpenReader(path);
            Assert.Equal(1, data.FieldCount);
            Assert.Equal("Value", data.GetName(0));
            Assert.True(data.Read());
            if (scenario == "text-colon") Assert.Equal("12:34", data.GetString(0));
            else Assert.Equal(42, data.GetInt32(0));
            Assert.True(data.Read());
            Assert.Equal(43, data.GetInt32(0));
            Assert.False(data.Read());
        } finally {
            File.Delete(path);
        }
    }
}
