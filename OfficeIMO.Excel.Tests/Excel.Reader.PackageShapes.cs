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

        private static byte[] CreateWorkbook(string rows, string? sharedStrings = null, string prefix = "") {
            var types = new List<(string, string)> {
                ("/xl/workbook.xml", WorkbookContentType),
                ("/xl/worksheets/sheet1.xml", WorksheetContentType)
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
            entries.Add(("xl/workbook.xml", Workbook(("Data", "rId1"))));
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
