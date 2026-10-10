using System.Data.Common;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("native", false)]
        [InlineData("native", true)]
        [InlineData("sdk", true)]
        [InlineData("sheet", false)]
        public void DataReader_XmlStructureRejectsNestedMarkupWithinInlineText(string surface, bool richText) {
            const string malformedText = "<t>A<foo>B</foo></t>";
            string content = richText ? "<r>" + malformedText + "</r>" : malformedText;
            string path = CreateXmlTextBudgetWorkbook(
                "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is>" + content + "</is></c></row>",
                multipleSheets: surface == "sdk");
            try {
                using var owner = surface == "sheet" ? ExcelDocumentReader.Open(path) : null;
                using DbDataReader reader = owner == null ? ExcelDocument.OpenDataReader(path)
                    : (DbDataReader)owner.GetSheet("Data").ReadRangeAsDataReader("A1:A4097", schemaSampleRows: 0);
                Assert.True(reader.Read());
                Assert.Throws<XmlException>(() => reader.GetString(0));
                reader.Close();
                Assert.True(reader.IsClosed);
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("native", true)]
        [InlineData("native", false)]
        [InlineData("sdk", true)]
        [InlineData("sheet", false)]
        public void DataReader_XmlStructureIgnoresCellExtensionValuesAndFormulas(string surface, bool useCachedFormulaResult) {
            string path = CreateXmlStructureExtensionWorkbook(surface == "sdk");
            try {
                var options = new ExcelReadOptions {
                    UseCachedFormulaResult = useCachedFormulaResult,
                    MaxXmlDataReaderBufferedCharacters = 32
                };
                using var owner = surface == "sheet" ? ExcelDocumentReader.Open(path, options) : null;
                using DbDataReader reader = owner == null ? ExcelDocument.OpenDataReader(path, options)
                    : (DbDataReader)owner.GetSheet("Data").ReadRangeAsDataReader("A1:D4097", schemaSampleRows: 0);
                Assert.True(reader.Read());
                Assert.Equal(1, reader.GetInt32(0));
                Assert.Equal("kept", reader.GetString(1));
                Assert.Equal("AB", reader.GetString(2));
                object expectedFormula = useCachedFormulaResult ? (object)3D : "1+2";
                Assert.Equal(expectedFormula, reader.GetValue(3));
                var values = new object[4];
                Assert.Equal(4, reader.GetValues(values));
                Assert.Equal(new object[] { 1D, "kept", "AB", expectedFormula }, values);
                Assert.True(reader.Read());
                Assert.Equal(2, reader.GetInt32(0));
                Assert.True(reader.IsDBNull(1));
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Reader_XmlStructureIgnoresExtensionsAcrossValueAndTypedObjectReaders(bool useConverter) {
            string path = CreateXmlStructureExtensionWorkbook();
            try {
                var options = new ExcelReadOptions {
                    CellValueConverter = useConverter ? _ => ExcelCellValue.NotHandled : null,
                    InferDataTableColumnTypes = false
                };
                using var owner = ExcelDocumentReader.Open(path, options);
                var sheet = owner.GetSheet("Data");
                object[] expected = { 1D, "kept", "AB", 3D };
                var values = sheet.ReadRange("A2:D2", ExcelExecutionMode.Sequential);
                for (int column = 0; column < expected.Length; column++) Assert.Equal(expected[column], values[0, column]);
                using var table = sheet.ReadRangeAsDataTable("A1:D2", headersInFirstRow: true);
                Assert.Single(table.Rows.Cast<System.Data.DataRow>());
                for (int column = 0; column < expected.Length; column++) Assert.Equal(expected[column], table.Rows[0][column]);
                var row = Assert.Single(sheet.ReadObjects<XmlStructureRecord>("A1:D2"));
                Assert.Equal(1D, row.Number);
                Assert.Equal("kept", row.Text);
                Assert.Equal("AB", row.Inline);
                Assert.Equal(3D, row.Formula);
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        [InlineData("http://purl.oclc.org/ooxml/spreadsheetml/main")]
        public void DataReader_XmlStructurePreservesVisibleRichTextAndDecodedCharacterLimit(string spreadsheetNamespace) {
            const string expected = " A<&B&🐢 \t ";
            string path = CreateCompactFastPathWorkbook();
            try {
                string xml = "<?xml version=\"1.0\" encoding=\"utf-16\"?>"
                    + "<worksheet xmlns=\"" + spreadsheetNamespace + "\" xmlns:s=\"" + spreadsheetNamespace + "\" xmlns:x=\"urn:extension\">"
                    + "<dimension ref=\"A1:A4097\"/><sheetData>"
                    + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Text</t></is></c></row>"
                    + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is>"
                    + "<r><rPr><b/></rPr><t xml:space=\"preserve\"> A</t></r>"
                    + "<rPh sb=\"0\" eb=\"1\"><t>phonetic</t></rPh>"
                    + "<x:r><t>extension</t></x:r><x:t>foreign</x:t>"
                    + "<r><s:t><![CDATA[<&]]>B&amp;🐢</s:t></r>"
                    + "<r><t xml:space=\"preserve\"> \t </t></r></is></c></row>"
                    + "<row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is/></c></row>"
                    + "<row r=\"4097\"><c r=\"A4097\"><v>0</v></c></row></sheetData></worksheet>";
                ReplaceZipEntry(path, "xl/worksheets/sheet1.xml",
                    Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray());
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = expected.Length })) {
                    Assert.True(reader.Read());
                    Assert.Equal(expected, reader.GetString(0));
                    Assert.Equal(expected, reader.GetValue(0));
                    Assert.True(reader.Read());
                    Assert.Equal(string.Empty, reader.GetString(0));
                }
                using var owner = ExcelDocumentReader.Open(path);
                Assert.Equal(expected, owner.GetSheet("Data").ReadRange("A2:A2", ExcelExecutionMode.Sequential)[0, 0]);
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = expected.Length - 1 })) {
                    Assert.True(reader.Read());
                    AssertXmlTextBudgetFailure(() => reader.GetString(0));
                    reader.Close();
                    Assert.True(reader.IsClosed);
                }
            } finally {
                File.Delete(path);
            }
        }

        private static string CreateXmlStructureExtensionWorkbook(bool multipleSheets = false) {
            const string headers = "<c r=\"A1\" t=\"inlineStr\"><is><t>Number</t></is></c>"
                + "<c r=\"B1\" t=\"inlineStr\"><is><t>Text</t></is></c>"
                + "<c r=\"C1\" t=\"inlineStr\"><is><t>Inline</t></is></c>"
                + "<c r=\"D1\" t=\"inlineStr\"><is><t>Formula</t></is></c>";
            const string rows = "<row r=\"2\" xmlns:x=\"urn:extension\">"
                + "<c r=\"A2\"><extLst><ext uri=\"number-before\"><v>888</v><x:v>777</x:v></ext></extLst>"
                + "<x:v>666</x:v><v>1</v><extLst><ext uri=\"number-after\"><v>555</v><x:v>999</x:v></ext></extLst></c>"
                + "<c r=\"B2\" t=\"str\"><v>kept</v><extLst><ext uri=\"text\">"
                + "<x:v>lost</x:v><x:f>extension-formula</x:f></ext></extLst><x:v>foreign</x:v></c>"
                + "<c r=\"C2\" t=\"inlineStr\"><is><r><t>A</t></r><r><t>B</t></r></is>"
                + "<extLst><ext uri=\"inline\"><is><t>nested</t></is><x:is><x:t>foreign</x:t></x:is></ext></extLst></c>"
                + "<c r=\"D2\"><f>1+2</f><v>3</v><extLst><ext uri=\"formula\">"
                + "<x:f>other</x:f><x:v>999</x:v></ext></extLst></c></row>"
                + "<row r=\"3\"><c r=\"A3\"><v>2</v></c></row>";
            return CreateXmlTextBudgetWorkbook(rows, columns: 4, headerCells: headers, multipleSheets: multipleSheets);
        }

        private sealed class XmlStructureRecord {
            public double Number { get; set; }
            public string? Text { get; set; }
            public string? Inline { get; set; }
            public object? Formula { get; set; }
        }
    }
}
