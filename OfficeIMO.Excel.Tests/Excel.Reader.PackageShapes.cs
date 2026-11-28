using System.IO.Compression;
using System.Text;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Cell values the forward-only reader returns for valid but less common package and cell
    /// markup, each built as the smallest package that shows the shape.
    /// </summary>
    public class ExcelReaderPackageShapeTests {
        private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        private const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

        [Theory]
        [InlineData("", "<c r=\"A1\" t=\"s\"/>")]
        [InlineData("", "<c r=\"A1\" t=\"s\"></c>")]
        [InlineData("x:", "<x:c r=\"A1\" t=\"s\"/>")]
        public void SharedStringCellWithoutValueIsEmpty(string prefix, string cell) {
            byte[] package = CreateWorkbook(
                $"<{prefix}row r=\"1\">{cell}<{prefix}c r=\"B1\" t=\"s\"><{prefix}v>0</{prefix}v></{prefix}c></{prefix}row>",
                sharedStrings: "<si><t>text</t></si>",
                prefix: prefix);

            Assert.Equal(new object?[] { null, "text" }, ReadFirstRow(package));
        }

        [Theory]
        [InlineData("true", true)]
        [InlineData("TRUE", true)]
        [InlineData(" true ", true)]
        [InlineData("false", false)]
        [InlineData("0", false)]
        public void BooleanCellsAcceptXsdBooleanLexicalForms(string text, bool expected) {
            byte[] package = CreateWorkbook($"<row r=\"1\"><c r=\"A1\" t=\"b\"><v>{text}</v></c></row>");

            Assert.Equal(new object?[] { expected }, ReadFirstRow(package));
        }

        [Fact]
        public void TimeOnlyIsoCellKeepsTickPrecision() {
            byte[] package = CreateWorkbook("<row r=\"1\"><c r=\"A1\" t=\"d\"><v>10:59:59.9999999999984025</v></c></row>");

            Assert.Equal(new object?[] { new DateTime(1899, 12, 30, 10, 59, 59).AddTicks(9_999_999) }, ReadFirstRow(package));
        }

        [Theory]
        [InlineData("http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet", "application/vnd.ms-excel.macrosheet+xml")]
        [InlineData("http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet", "application/vnd.ms-excel.intlmacrosheet+xml")]
        public void MacroSheetsAreSkippedLikeOtherNonWorksheetSheets(string relationshipType, string contentType) {
            byte[] package = CreatePackage(
                ("[Content_Types].xml", ContentTypes(
                    ("/xl/workbook.xml", WorkbookContentType),
                    ("/xl/macrosheets/sheet1.xml", contentType),
                    ("/xl/worksheets/sheet1.xml", WorksheetContentType))),
                ("_rels/.rels", Relationships(("rId1", "officeDocument", "xl/workbook.xml"))),
                ("xl/workbook.xml", Workbook(("Macro1", "rId1"), ("Data", "rId2"))),
                ("xl/_rels/workbook.xml.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + $"<Relationship Id=\"rId1\" Type=\"{relationshipType}\" Target=\"macrosheets/sheet1.xml\"/>"
                    + $"<Relationship Id=\"rId2\" Type=\"{RelationshipNamespace}/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                    + "</Relationships>"),
                ("xl/macrosheets/sheet1.xml",
                    "<xm:macrosheet xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\" "
                    + $"xmlns=\"{SpreadsheetNamespace}\"><sheetData/></xm:macrosheet>"),
                ("xl/worksheets/sheet1.xml", Worksheet("<row r=\"1\"><c r=\"A1\"><v>7</v></c></row>")));

            using var reader = ExcelDocument.OpenReader(new MemoryStream(package), Options());
            Assert.True(reader.Read());
            Assert.Equal(7d, reader.GetValue(0));
            Assert.False(reader.NextResult());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void WorkbookAndCellLookupsIgnoreExtensionAndAlternateContent(bool wrapLiveWorkbook) {
            string workbook = $"""
                <workbook xmlns="{SpreadsheetNamespace}" xmlns:r="{RelationshipNamespace}"
                          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                          xmlns:vendor="urn:extension">
                  <extLst><ext uri="urn:inventory"><sheets><sheet name="Extension" sheetId="9" r:id="rId1"/></sheets></ext></extLst>
                  <mc:AlternateContent>
                    <mc:Choice Requires="vendor"><sheets><sheet name="Choice" sheetId="8" r:id="rId1"/></sheets></mc:Choice>
                    <mc:Fallback><sheets><sheet name="Fallback" sheetId="7" r:id="rId1"/></sheets></mc:Fallback>
                  </mc:AlternateContent>
                  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """;
            byte[] package = CreateWorkbook("""
                <row r="1"><c r="A1"><extLst><ext uri="urn:cell"><v>999</v></ext></extLst><v>7</v></c></row>
                """, workbookMarkup: workbook);

            using var stream = new MemoryStream(package);
            using var live = wrapLiveWorkbook ? DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(stream, true,
                new DocumentFormat.OpenXml.Packaging.OpenSettings { AutoSave = false }) : null;
            using var reader = live == null ? ExcelDocumentReader.Open(package) : ExcelDocumentReader.Wrap(live);
            Assert.Equal(new[] { "Data" }, reader.GetValidatedWorksheetNames());
            Assert.Equal("Data", reader.GetSheet(1).Name);
            Assert.Equal(7d, Assert.Single(reader.GetSheet("Data").EnumerateCells()).Value);
        }

        [Fact]
        public void StandaloneReaderResolvesSharedStringsByExactRelationshipType() {
            byte[] package = CreatePackage(
                ("[Content_Types].xml", ContentTypes(
                    ("/xl/workbook.xml", WorkbookContentType), ("/xl/worksheets/sheet1.xml", WorksheetContentType),
                    ("/xl/foreign.xml", SharedStringsContentType), ("/xl/sharedStrings.xml", SharedStringsContentType))),
                ("_rels/.rels", Relationships(("rId1", "officeDocument", "xl/workbook.xml"))),
                ("xl/workbook.xml", Workbook(("Data", "rId1"))),
                ("xl/_rels/workbook.xml.rels", $"""
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rId1" Type="{RelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
                      <Relationship Id="rId2" Type="https://example.org/workbook/relationships/sharedStrings" Target="foreign.xml"/>
                      <Relationship Id="rId3" Type="{RelationshipNamespace}/sharedStrings" Target="sharedStrings.xml"/>
                    </Relationships>
                    """),
                ("xl/worksheets/sheet1.xml", Worksheet("""<row r="1"><c r="A1" t="s"><v>0</v></c></row>""")),
                ("xl/foreign.xml", SharedStrings("<si><t>unrelated</t></si>")),
                ("xl/sharedStrings.xml", SharedStrings("<si><t>actual</t></si>")));

            using var reader = ExcelDocumentReader.Open(package);
            Assert.Equal("actual", Assert.Single(reader.GetSheet(1).EnumerateCells()).Value);
        }

        [Fact]
        public void StandaloneReaderRejectsWorksheetRelationshipWithTheWrongContentType() {
            byte[] package = CreateWorkbook("""<row r="1"><c r="A1"><v>7</v></c></row>""",
                worksheetContentType: "application/xml");
            using var reader = ExcelDocumentReader.Open(package);
            IOException error = Assert.Throws<IOException>(() => reader.GetValidatedWorksheetNames());
            Assert.IsType<InvalidDataException>(error.InnerException);
        }

        [Fact]
        public void CompressedWorkbookInputLimitDoesNotApplyToInflatedWorksheet() {
            var options = Options();
            options.Culture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL");
            options.MaxInputBytes = 4096;
            options.A1Range = "A1:A2";
            byte[] package = CreateWorkbook(
                """<row r="1"><c r="A1"><v>7</v></c></row><row r="2"><c r="A2"><v>8</v></c></row>"""
                + "<!-- " + new string('x', 32_768) + " -->");
            Assert.True(package.Length < options.MaxInputBytes);

            using var reader = ExcelDocument.OpenReader(new MemoryStream(package), options);
            Assert.True(reader.Read());
            Assert.Equal(7d, reader.GetDouble(0));
            Assert.True(reader.Read());
            Assert.Equal(8d, reader.GetDouble(0));
            Assert.False(reader.Read());
        }

        private static ExcelReaderOptions Options() => new() {
            HasHeaderRow = false,
            TreatDatesUsingNumberFormat = false
        };

        private static object?[] ReadFirstRow(byte[] package) {
            using var reader = ExcelDocument.OpenReader(new MemoryStream(package), Options());
            Assert.True(reader.Read());
            var values = new object?[reader.FieldCount];
            for (int ordinal = 0; ordinal < values.Length; ordinal++) {
                values[ordinal] = reader.IsDBNull(ordinal) ? null : reader.GetValue(ordinal);
            }
            return values;
        }

        private static byte[] CreateWorkbook(string rows, string? sharedStrings = null, string prefix = "", string? workbookMarkup = null, string worksheetContentType = WorksheetContentType) {
            var types = new List<(string, string)> {
                ("/xl/workbook.xml", WorkbookContentType),
                ("/xl/worksheets/sheet1.xml", worksheetContentType)
            };
            var relationships = new List<(string, string, string)> { ("rId1", "worksheet", "worksheets/sheet1.xml") };
            var entries = new List<(string, string)>();
            if (sharedStrings != null) {
                types.Add(("/xl/sharedStrings.xml", SharedStringsContentType));
                relationships.Add(("rId2", "sharedStrings", "sharedStrings.xml"));
                entries.Add(("xl/sharedStrings.xml", SharedStrings(sharedStrings)));
            }

            entries.Add(("[Content_Types].xml", ContentTypes(types.ToArray())));
            entries.Add(("_rels/.rels", Relationships(("rId1", "officeDocument", "xl/workbook.xml"))));
            entries.Add(("xl/workbook.xml", workbookMarkup ?? Workbook(("Data", "rId1"))));
            entries.Add(("xl/_rels/workbook.xml.rels", Relationships(relationships.ToArray())));
            entries.Add(("xl/worksheets/sheet1.xml", prefix.Length == 0
                ? Worksheet(rows)
                : $"<{prefix}worksheet xmlns:{prefix.TrimEnd(':')}=\"{SpreadsheetNamespace}\"><{prefix}sheetData>{rows}</{prefix}sheetData></{prefix}worksheet>"));
            return CreatePackage(entries.ToArray());
        }

        private static string ContentTypes(params (string PartName, string ContentType)[] overrides) =>
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + string.Concat(overrides.Select(item => $"<Override PartName=\"{item.PartName}\" ContentType=\"{item.ContentType}\"/>"))
            + "</Types>";

        private static string Relationships(params (string Id, string Type, string Target)[] relationships) =>
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + string.Concat(relationships.Select(item =>
                $"<Relationship Id=\"{item.Id}\" Type=\"{RelationshipNamespace}/{item.Type}\" Target=\"{item.Target}\"/>"))
            + "</Relationships>";

        private static string Workbook(params (string Name, string RelationshipId)[] sheets) =>
            $"<workbook xmlns=\"{SpreadsheetNamespace}\" xmlns:r=\"{RelationshipNamespace}\"><sheets>"
            + string.Concat(sheets.Select((sheet, index) =>
                $"<sheet name=\"{sheet.Name}\" sheetId=\"{index + 1}\" r:id=\"{sheet.RelationshipId}\"/>"))
            + "</sheets></workbook>";

        private static string Worksheet(string rows) =>
            $"<worksheet xmlns=\"{SpreadsheetNamespace}\"><sheetData>{rows}</sheetData></worksheet>";

        private static string SharedStrings(string items) =>
            $"<sst xmlns=\"{SpreadsheetNamespace}\">{items}</sst>";

        private static byte[] CreatePackage(params (string Name, string Content)[] entries) {
            using var output = new MemoryStream();
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                foreach ((string name, string content) in entries) {
                    using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
                    writer.Write(content);
                }
            }
            return output.ToArray();
        }
    }
}
