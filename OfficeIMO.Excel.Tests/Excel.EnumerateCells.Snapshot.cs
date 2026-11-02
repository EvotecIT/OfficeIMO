using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Cell enumeration of an unchanged workbook reads the package bytes it was loaded from, so
    /// these cover the boundary with editing and the values only that route has to reproduce.
    /// </summary>
    public class ExcelEnumerateCellsSnapshotTests {
        private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        [Fact]
        public void EnumerationFollowsEditsMadeBetweenAndDuringReads() {
            byte[] package = CreatePackage("<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row>");
            using var document = ExcelDocument.Load(new MemoryStream(package));
            ExcelSheet sheet = document.Sheets[0];
            Assert.Equal(new object[] { 1d, 2d }, sheet.EnumerateCells().Select(cell => cell.Value!).ToArray());

            using (var cells = sheet.EnumerateCells().GetEnumerator()) {
                Assert.True(cells.MoveNext());
                sheet.CellValue(1, 1, 10d);
                Assert.True(cells.MoveNext());
                Assert.Equal(2d, cells.Current.Value);
                Assert.False(cells.MoveNext());
            }

            Assert.Equal(new object[] { 10d, 2d }, sheet.EnumerateCells().Select(cell => cell.Value!).ToArray());

            using var saved = new MemoryStream();
            document.Save(saved);
            Assert.Equal(new object[] { 10d, 2d }, sheet.EnumerateCells().Select(cell => cell.Value!).ToArray());

            sheet.CellValue(1, 2, 20d);
            Assert.Equal(new object[] { 10d, 20d }, sheet.EnumerateCells().Select(cell => cell.Value!).ToArray());
        }

        [Fact]
        public void EditedWorkbookReadsUntouchedSheetsFromThePackageWithoutLoadingThem() {
            byte[] package = CreatePackage(
                ("Edited", "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>"),
                ("Untouched", "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\"><v>2</v></c></row>"));
            using var document = ExcelDocument.Load(new MemoryStream(package));
            ExcelSheet edited = document["Edited"];
            ExcelSheet untouched = document["Untouched"];
            edited.CellValue(1, 2, "added");

            using (var cells = untouched.EnumerateCells().GetEnumerator()) {
                Assert.True(cells.MoveNext());
                Assert.Equal("shared", cells.Current.Value);
                edited.CellValue(1, 3, "added later");
                Assert.True(cells.MoveNext());
                Assert.Equal(2d, cells.Current.Value);
            }

            Assert.Equal(new object[] { "shared", "added", "added later" }, edited.EnumerateCells().Select(cell => cell.Value!).ToArray());
            Assert.Equal("shared", untouched.ToDataTable(headersInFirstRow: false).Rows[0][0]);

            var parts = document.OpenXmlDocument.WorkbookPart!;
            var loaded = parts.Workbook!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                .ToDictionary(sheet => sheet.Name!.Value!, sheet => parts.GetPartById(sheet.Id!.Value!).IsRootElementLoaded);
            Assert.True(loaded["Edited"]);
            Assert.False(loaded["Untouched"]);
        }

        [Fact]
        public void EnumerationSeesChangesMadeThroughTheExposedSdkDocument() {
            byte[] package = CreatePackage("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>");
            using var document = ExcelDocument.Load(new MemoryStream(package));
            ExcelSheet sheet = document.Sheets[0];
            Assert.Equal(1d, sheet.EnumerateCells().Single().Value);

            var worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.CellValue>().Single().Text = "5";

            Assert.Equal(5d, sheet.EnumerateCells().Single().Value);
        }

        [Fact]
        public void ExplicitBlankAndWhitespaceValuesAreCellContent() {
            byte[] package = CreatePackage(
                "<row r=\"1\"><c r=\"A1\"><v></v></c><c r=\"B1\" t=\"str\"><f>\"\"</f><v></v></c>"
                + "<c r=\"C1\" t=\"str\"><v>  </v></c><c r=\"D1\" t=\"inlineStr\"><is><t>  </t></is></c>"
                + "<c r=\"E1\" t=\"s\"><v>1</v></c><c r=\"F1\"/></row>");
            using var document = ExcelDocument.Load(new MemoryStream(package));
            ExcelSheet sheet = document.Sheets[0];

            var cells = sheet.EnumerateCells().Select(cell => (cell.Column, cell.Value)).ToArray();
            Assert.Equal(new (int, object?)[] { (1, ""), (2, ""), (3, "  "), (4, "  "), (5, "   ") }, cells);

            var range = sheet.EnumerateRange("A1:E1").Select(cell => cell.Value).ToArray();
            Assert.Equal(new object?[] { "", "", "  ", "  ", "   " }, range);
        }

        [Fact]
        public void InvalidCellTypeIsRejectedWhetherOrNotTheWorkbookHasBeenEdited() {
            byte[] package = CreatePackage("<row r=\"1\"><c r=\"A1\" t=\"zz\"><v>1</v></c></row>");
            using var document = ExcelDocument.Load(new MemoryStream(package));

            Assert.Throws<FormatException>(() => document.Sheets[0].EnumerateCells().ToArray());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void OpenDocumentKeepsItsContentWhenTheSourceFileIsReplaced(bool fileBacked) {
            string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Snapshot.{Guid.NewGuid():N}.xlsx");
            try {
                File.WriteAllBytes(path, CreatePackage("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"));
                using var document = ExcelDocument.Load(path, new ExcelLoadOptions { FileBacked = fileBacked });
                ExcelSheet sheet = document.Sheets[0];

                File.WriteAllBytes(path, CreatePackage("<row r=\"1\"><c r=\"A1\"><v>2</v></c></row>"));

                Assert.Equal(1d, sheet.EnumerateCells().Single().Value);
                using var reader = sheet.CreateDataReader(new ExcelReaderOptions { HasHeaderRow = false });
                Assert.True(reader.Read());
                Assert.Equal(1d, reader.GetValue(0));
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void FileBackedEnumerationSurvivesSaveAndFollowsTheSavedWorkbook() {
            string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.Snapshot.{Guid.NewGuid():N}.xlsx");
            try {
                File.WriteAllBytes(path, CreatePackage("<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row>"));
                using var document = ExcelDocument.Load(path, new ExcelLoadOptions { FileBacked = true });
                ExcelSheet sheet = document.Sheets[0];

                using (var cells = sheet.EnumerateCells().GetEnumerator()) {
                    Assert.True(cells.MoveNext());
                    sheet.CellValue(1, 1, 10d);
                    document.Save();
                    Assert.True(cells.MoveNext());
                    Assert.Equal(2d, cells.Current.Value);
                }

                Assert.Equal(new object[] { 10d, 2d }, sheet.EnumerateCells().Select(cell => cell.Value!).ToArray());
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void TabularReadersKeepTheSdkReadingOfNonCanonicalWorksheets() {
            byte[] whitespace = CreatePackage("<row r=\"1\"><c r=\"A1\" t=\"str\"><v>   </v></c></row>");
            using (var document = ExcelDocument.Load(new MemoryStream(whitespace))) {
                Assert.Equal("   ", document.Sheets[0].ToDataTable(headersInFirstRow: false).Rows[0][0]);
            }

            byte[] outsideGrid = CreatePackage(
                "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"1048577\"><c r=\"A1048577\"><v>2</v></c></row>");
            using (var document = ExcelDocument.Load(new MemoryStream(outsideGrid))) {
                using var reader = document.Sheets[0].CreateDataReader(new ExcelReaderOptions { HasHeaderRow = false });
                Assert.True(reader.Read());
                Assert.Equal(1d, reader.GetValue(0));
            }
        }

        [Fact]
        public void NonCanonicalWorksheetKeepsTheSdkRouteWithoutAffectingItsNeighbour() {
            byte[] package = CreatePackage(
                ("Good", "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"),
                ("Bad", "<row r=\"1\"><c r=\"A1\" t=\"str\"><v>   </v></c></row>"));
            using var document = ExcelDocument.Load(new MemoryStream(package));

            Assert.Equal(1d, document["Good"].ToDataTable(headersInFirstRow: false).Rows[0][0]);
            Assert.Equal("   ", document["Bad"].ToDataTable(headersInFirstRow: false).Rows[0][0]);

            var parts = document.OpenXmlDocument.WorkbookPart!;
            var loaded = parts.Workbook!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                .ToDictionary(sheet => sheet.Name!.Value!, sheet => parts.GetPartById(sheet.Id!.Value!).IsRootElementLoaded);
            Assert.False(loaded["Good"]);
            Assert.True(loaded["Bad"]);
        }

        [Fact]
        public void StreamingReadersPassWholeSheetsThatDenseReadsStillLimit() {
            byte[] package = CreatePackage(string.Concat(Enumerable.Range(1, 30).Select(row =>
                $"<row r=\"{row}\"><c r=\"A{row}\"><v>{row}</v></c><c r=\"B{row}\"><v>{row}</v></c></row>")));
            using var document = ExcelDocument.Load(new MemoryStream(package));
            ExcelSheet sheet = document.Sheets[0];
            var options = new ExcelReaderOptions { MaxRangeCells = 10, HasHeaderRow = false };

            using (var reader = sheet.CreateDataReader(options)) {
                int rows = 0;
                while (reader.Read()) rows++;
                Assert.Equal(30, rows);
            }

            Assert.Throws<OfficeIMO.OfficeInspectionResourceLimitException>(() => sheet.ToDataTable(false, options));
        }

        [Fact]
        public void ChunkOverTheBufferedCellLimitFailsTheSameOnBothRoutes() {
            byte[] package = CreatePackage(
                "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"1048576\"><c r=\"XFD1048576\"><v>2</v></c></row>");
            var options = new ExcelReaderOptions { HasHeaderRow = false };
            static void ReadFirstRow(ExcelSheet sheet, ExcelReaderOptions readerOptions) {
                using var reader = sheet.CreateDataReader(readerOptions);
                reader.Read();
            }

            using var loaded = ExcelDocument.Load(new MemoryStream(package));
            using var exposed = ExcelDocument.Load(new MemoryStream(package));
            _ = exposed.OpenXmlDocument;

            Assert.Throws<OfficeIMO.OfficeInspectionResourceLimitException>(() => ReadFirstRow(exposed.Sheets[0], options));
            Assert.Throws<OfficeIMO.OfficeInspectionResourceLimitException>(() => ReadFirstRow(loaded.Sheets[0], options));
        }

        [Fact]
        public void SheetWiderThanOneMaximumChunkStaysOnTheReadOnlyRoute() {
            byte[] package = CreatePackage(string.Concat(Enumerable.Range(1, 30).Select(row =>
                $"<row r=\"{row}\">{string.Concat("ABCDE".Select(column => $"<c r=\"{column}{row}\"><v>{row}</v></c>"))}</row>")));
            var options = new ExcelReaderOptions {
                HasHeaderRow = false, MaxDataReaderBufferedCells = 50, MaxDataReaderChunkRows = 8, SchemaSampleRows = 1, MaxDataReaderSchemaSampleRows = 20
            };
            using var document = ExcelDocument.Load(new MemoryStream(package));

            int cells = 0;
            using (var reader = document.Sheets[0].CreateDataReader(options)) {
                while (reader.Read()) cells += reader.FieldCount;
            }

            Assert.Equal(150, cells);
            Assert.False(document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().IsRootElementLoaded);
        }

        [Fact]
        public void CultureAppliesToNumberText() {
            byte[] package = CreatePackage("<row r=\"1\"><c r=\"A1\"><v>1,25</v></c></row>");
            using var document = ExcelDocument.Load(new MemoryStream(package));
            var options = new ExcelReaderOptions { Culture = System.Globalization.CultureInfo.GetCultureInfo("de-DE") };

            Assert.Equal(1.25d, document.Sheets[0].EnumerateCells(options).Single().Value);
        }

        [Fact]
        public void CellValueConverterSeesCellsWithoutContent() {
            byte[] package = CreatePackage("<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"/></row>");
            using var document = ExcelDocument.Load(new MemoryStream(package));
            var options = new ExcelReaderOptions { CellValueConverter = context => new ExcelCellValue(context.RawText ?? "empty") };

            object?[] values = document.Sheets[0].EnumerateCells(options).Select(cell => cell.Value).ToArray();

            Assert.Equal(new object?[] { "1", "empty" }, values);
        }

        [Fact]
        public void CultureAppliesToDateTextCells() {
            byte[] package = CreatePackage("<row r=\"1\"><c r=\"A1\" t=\"d\"><v>01.02.2024</v></c></row>");
            using var document = ExcelDocument.Load(new MemoryStream(package));
            var options = new ExcelReaderOptions { Culture = System.Globalization.CultureInfo.GetCultureInfo("de-DE") };

            Assert.Equal(new DateTime(2024, 2, 1), document.Sheets[0].EnumerateCells(options).Single().Value);
        }

        private static byte[] CreatePackage(string rows) => CreatePackage(("Data", rows));

        private static byte[] CreatePackage(params (string Name, string Rows)[] sheets) {
            var overrides = new StringBuilder();
            var sheetElements = new StringBuilder();
            var relationships = new StringBuilder();
            for (int index = 1; index <= sheets.Length; index++) {
                overrides.Append($"<Override PartName=\"/xl/worksheets/sheet{index}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                sheetElements.Append($"<sheet name=\"{sheets[index - 1].Name}\" sheetId=\"{index}\" r:id=\"rId{index}\"/>");
                relationships.Append($"<Relationship Id=\"rId{index}\" Type=\"{RelationshipNamespace}/worksheet\" Target=\"worksheets/sheet{index}.xml\"/>");
            }

            using var output = new MemoryStream();
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                Add(archive, "[Content_Types].xml",
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                    + overrides
                    + "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
                    + "</Types>");
                Add(archive, "_rels/.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + $"<Relationship Id=\"rId1\" Type=\"{RelationshipNamespace}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                Add(archive, "xl/workbook.xml",
                    $"<workbook xmlns=\"{SpreadsheetNamespace}\" xmlns:r=\"{RelationshipNamespace}\"><sheets>{sheetElements}</sheets></workbook>");
                Add(archive, "xl/_rels/workbook.xml.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + relationships
                    + $"<Relationship Id=\"rId900\" Type=\"{RelationshipNamespace}/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>");
                for (int index = 1; index <= sheets.Length; index++) {
                    Add(archive, $"xl/worksheets/sheet{index}.xml",
                        $"<worksheet xmlns=\"{SpreadsheetNamespace}\"><sheetData>{sheets[index - 1].Rows}</sheetData></worksheet>");
                }
                Add(archive, "xl/sharedStrings.xml",
                    $"<sst xmlns=\"{SpreadsheetNamespace}\" count=\"2\" uniqueCount=\"2\"><si><t>shared</t></si><si><t>   </t></si></sst>");
            }

            return output.ToArray();
        }

        private static void Add(ZipArchive archive, string name, string content) {
            using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
            writer.Write(content);
        }
    }
}
