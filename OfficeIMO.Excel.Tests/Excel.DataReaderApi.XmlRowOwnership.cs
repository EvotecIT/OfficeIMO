using System.Data.Common;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("native")]
        [InlineData("sdk")]
        [InlineData("sheet")]
        public void DataReader_XmlRowOwnershipIgnoresExtensionCells(string surface) {
            string path = CreateXmlRowOwnershipWorkbook(multipleSheets: surface == "sdk");
            try {
                var options = new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = 11 };
                using var owner = surface == "sheet" ? ExcelDocumentReader.Open(path, options) : null;
                using DbDataReader reader = owner == null ? ExcelDocument.OpenDataReader(path, options)
                    : (DbDataReader)owner.GetSheet("Data").ReadRangeAsDataReader("A1:B4097", schemaSampleRows: 0);
                Assert.Equal(2, reader.FieldCount);
                Assert.Equal("Number", reader.GetName(0));
                Assert.Equal("Other", reader.GetName(1));
                Assert.True(reader.Read());
                Assert.Equal(1, reader.GetInt32(0));
                Assert.True(reader.IsDBNull(1));
                var values = new object[2];
                Assert.Equal(2, reader.GetValues(values));
                Assert.Equal(new object[] { 1D, DBNull.Value }, values);
                Assert.True(reader.Read());
                Assert.Equal(2, reader.GetInt32(0));
                Assert.Equal(3, reader.GetInt32(1));
                reader.Close();
                Assert.True(reader.IsClosed);
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Sheet_XmlRowOwnershipIgnoresExtensionCellsInPublicProjections(bool useSdkControl) {
            string path = CreateXmlRowOwnershipWorkbook();
            try {
                var options = new ExcelReadOptions {
                    Culture = useSdkControl ? CultureInfo.GetCultureInfo("fr-FR") : CultureInfo.InvariantCulture,
                    InferDataTableColumnTypes = false
                };
                using var document = ExcelDocument.Load(path);
                var sheet = document.Sheets.Single();
                var cells = sheet.EnumerateRange("A2:B3", options).ToArray();
                Assert.Equal(new[] { (2, 1), (3, 1), (3, 2) }, cells.Select(cell => (cell.Row, cell.Column)).ToArray());
                Assert.Equal(new object?[] { 1D, 2D, 3D }, cells.Select(cell => cell.Value).ToArray());
                using var table = sheet.ToDataTable("A1:B3", options: options);
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal(1D, table.Rows[0][0]);
                Assert.True(table.Rows[0].IsNull(1));
                Assert.Equal(2D, table.Rows[1][0]);
                Assert.Equal(3D, table.Rows[1][1]);
                using var owner = ExcelDocumentReader.Open(path, options);
                var objects = owner.GetSheet("Data").ReadObjects<XmlRowOwnershipRecord>("A1:B3").ToArray();
                Assert.Equal(2, objects.Length);
                Assert.Equal(1D, objects[0].Number);
                Assert.Null(objects[0].Other);
                Assert.Equal(2D, objects[1].Number);
                Assert.Equal(3D, objects[1].Other);
            } finally {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        [InlineData("http://purl.oclc.org/ooxml/spreadsheetml/main")]
        public void Reader_XmlRowOwnershipPreservesActualBoundsAndImplicitCoordinates(string spreadsheetNamespace) {
            string path = CreateXmlRowOwnershipWorkbook(implicitCoordinates: true, spreadsheetNamespace: spreadsheetNamespace);
            try {
                using (DbDataReader reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { MaxXmlDataReaderBufferedCharacters = 11 })) {
                    Assert.Equal(2, reader.FieldCount);
                    Assert.True(reader.Read());
                    Assert.Equal(1, reader.GetInt32(0));
                    Assert.Equal(2, reader.GetInt32(1));
                    Assert.True(reader.Read());
                    Assert.Equal(3, reader.GetInt32(0));
                    Assert.Equal(4, reader.GetInt32(1));
                }

                using var owner = ExcelDocumentReader.Open(path, new ExcelReadOptions { InferDataTableColumnTypes = false });
                var sheet = owner.GetSheet("Data");
                Assert.Equal("A1:B4097", sheet.GetUsedRangeA1());
                foreach (var mode in new[] { ExcelExecutionMode.Sequential, ExcelExecutionMode.Parallel }) {
                    var values = sheet.ReadRange("A2:B3", mode);
                    Assert.Equal(1D, values[0, 0]);
                    Assert.Equal(2D, values[0, 1]);
                    Assert.Equal(3D, values[1, 0]);
                    Assert.Equal(4D, values[1, 1]);
                }
                var rows = sheet.ReadRows("A2:B3").ToArray();
                Assert.Equal(2, rows.Length);
                Assert.Equal(new object?[] { 1D, 2D }, rows[0]);
                Assert.Equal(new object?[] { 3D, 4D }, rows[1]);
                Assert.Equal(new object?[] { 2D, 4D }, sheet.ReadColumn("B2:B3").ToArray());
                var chunks = sheet.ReadRangeStream("A2:B3", chunkRows: 1).ToArray();
                Assert.Equal(2, chunks.Length);
                Assert.Equal(new object?[] { 1D, 2D }, chunks[0].Rows[0]);
                Assert.Equal(new object?[] { 3D, 4D }, chunks[1].Rows[0]);
                var objects = sheet.ReadObjectsStream<XmlRowOwnershipRecord>("A1:B3").ToArray();
                Assert.Equal(2, objects.Length);
                Assert.Equal(1D, objects[0].Number);
                Assert.Equal(2D, objects[0].Other);
                Assert.Equal(3D, objects[1].Number);
                Assert.Equal(4D, objects[1].Other);
            } finally {
                File.Delete(path);
            }
        }

        private static string CreateXmlRowOwnershipWorkbook(bool multipleSheets = false, bool implicitCoordinates = false,
            string spreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main") {
            string path = CreateXmlTextBudgetWorkbook(string.Empty, multipleSheets: multipleSheets);
            try {
                string extensionCells = "<s:extLst><s:ext uri=\"cells\">"
                    + "<x:c r=\"B2\"><s:v>999</s:v></x:c>"
                    + "<s:c r=\"B2\"><s:v>888</s:v></s:c>"
                    + "<x:c r=\"B2\" t=\"inlineStr\"><s:is><s:t>" + new string('x', 100) + "</s:t></s:is></x:c>"
                    + (implicitCoordinates ? "<s:c r=\"XFD900000\"><s:v>999</s:v></s:c>" : string.Empty)
                    + "</s:ext></s:extLst><x:c r=\"B2\"><s:v>777</s:v></x:c>";
                string rows = implicitCoordinates
                    ? "<s:row>" + extensionCells + "<s:c r=\"A2\"><s:v>1</s:v></s:c><s:c><s:v>2</s:v></s:c></s:row>"
                        + "<s:row><s:c><s:v>3</s:v></s:c><s:c><s:v>4</s:v></s:c></s:row>"
                    : "<s:row r=\"2\"><s:c r=\"A2\"><s:v>1</s:v></s:c>" + extensionCells + "</s:row>"
                        + "<s:row r=\"3\"><s:c r=\"A3\"><s:v>2</s:v></s:c><s:c r=\"B3\"><s:v>3</s:v></s:c></s:row>";
                // These same-namespace rows have the real row depth but belong to
                // foreign worksheet children. Depth and namespace alone are insufficient.
                string foreignRows = "<x:payload><s:row r=\"2\"><s:c r=\"B2\"><s:v>666</s:v></s:c>"
                    + "<s:c r=\"Z999999\"><s:v>999</s:v></s:c></s:row></x:payload>";
                string xml = "<?xml version=\"1.0\" encoding=\"utf-16\"?>"
                    + "<s:worksheet xmlns:s=\"" + spreadsheetNamespace + "\" xmlns:x=\"urn:extension\">"
                    + foreignRows + "<s:sheetData>"
                    + "<s:row r=\"1\"><s:c r=\"A1\" t=\"inlineStr\"><s:is><s:t>Number</s:t></s:is></s:c>"
                    + "<s:c r=\"B1\" t=\"inlineStr\"><s:is><s:t>Other</s:t></s:is></s:c></s:row>"
                    + "<x:row r=\"2\"><s:c r=\"B2\"><s:v>555</s:v></s:c></x:row>"
                    + rows + "<s:row r=\"4097\"><s:c r=\"A4097\"><s:v>0</s:v></s:c></s:row>"
                    + "</s:sheetData>" + foreignRows + "</s:worksheet>";
                // UTF16 exercises XML scanners; the real final row selects streaming.
                ReplaceZipEntry(path, "xl/worksheets/sheet1.xml",
                    Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(xml)).ToArray());
                return path;
            } catch {
                File.Delete(path);
                throw;
            }
        }

        private sealed class XmlRowOwnershipRecord {
            public double Number { get; set; }
            public double? Other { get; set; }
        }
    }
}
