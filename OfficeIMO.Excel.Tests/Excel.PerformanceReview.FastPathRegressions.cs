using System.Data;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void PerformanceReview_Utf8ObjectReaderDoesNotCoerceFormulaTextIntoTypedProperties(bool useCachedFormulaResult) {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions {
                PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
            })) {
                var sheet = document.AddSheet("Formula");
                sheet.Cell(2, 1).SetValue("Id");
                sheet.Cell(2, 2).SetValue("Active");
                sheet.Cell(2, 3).SetValue("Expression");
                sheet.Cell(3, 1).SetFormula("1");
                sheet.Cell(3, 2).SetFormula("TRUE");
                sheet.Cell(3, 3).SetFormula("TRUE");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray(), new ExcelReaderOptions {
                UseCachedFormulaResult = useCachedFormulaResult
            });
            FormulaProjection row = reader.GetSheet("Formula")
                .ReadObjectsStream<FormulaProjection>("A2:C3")
                .Single();

            Assert.Equal(0, row.Id);
            Assert.False(row.Active);
            Assert.Equal("TRUE", row.Expression);
        }

        [Fact]
        public void PerformanceReview_XmlDataReaderDoesNotReusePrimitiveValuesForMissingCells() {
            var expectedDate = new DateTime(2026, 7, 14);
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory, new ExcelCreateOptions {
                PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
            })) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Id");
                sheet.Cell(1, 2).SetValue("Created");
                sheet.Cell(1, 3).SetValue("Active");
                sheet.Cell(1, 4).SetValue("Note");
                sheet.Cell(2, 1).SetValue(42);
                sheet.Cell(2, 2).SetValue(expectedDate);
                sheet.Cell(2, 3).SetValue(true);
                sheet.Cell(2, 4).SetValue("Complete");
                sheet.Cell(3, 4).SetValue("Missing typed values");
            }

            byte[] utf16WorksheetPackage = RewriteFirstWorksheetAsUtf16(memory.ToArray());
            using var reader = ExcelDocumentReader.Open(utf16WorksheetPackage);
            using var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:D3", schemaSampleRows: 0);

            Assert.True(dataReader.Read());
            Assert.Equal(42, dataReader.GetInt32(0));
            Assert.Equal(expectedDate, dataReader.GetDateTime(1));
            Assert.True(dataReader.GetBoolean(2));

            Assert.True(dataReader.Read());
            Assert.Equal(DBNull.Value, dataReader.GetValue(0));
            Assert.Equal(DBNull.Value, dataReader.GetValue(1));
            Assert.Equal(DBNull.Value, dataReader.GetValue(2));
            Assert.Equal("Missing typed values", dataReader.GetString(3));
        }

        [Fact]
        public void PerformanceReview_ExtendedPackageStagesWritesWhenDestinationBacksRawParts() {
            using var associatedDestination = new MemoryStream();
            using (var document = ExcelDocument.Create(associatedDestination, new ExcelCreateOptions {
                PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
            })) {
                var rows = new[] {
                    new FastPackageProjection("Alpha", 10),
                    new FastPackageProjection("Beta", 20)
                };
                document.DefaultChartStylePreset = ExcelChartStylePreset.Default;
                var sheet = document.AddSheet("Data");
                sheet.WriteData(rows, ("Name", row => row.Name), ("Score", row => row.Score));
                sheet.AddChart(OfficeChartType.ColumnClustered, new ExcelChartData(
                        rows.Select(row => row.Name),
                        new[] { new ExcelChartSeries("Score", rows.Select(row => (double)row.Score)) }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5, Title = "Scores" });

                var packageStreamField = typeof(ExcelDocument).GetField("_packageStream", BindingFlags.Instance | BindingFlags.NonPublic);
                var packageStream = Assert.IsAssignableFrom<Stream>(packageStreamField!.GetValue(document));

                document.Save(packageStream);

                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
            }

            associatedDestination.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(associatedDestination, false);
            var dataSheet = spreadsheet.WorkbookPart!.Workbook.Sheets!
                .Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                .Single(sheet => string.Equals(sheet.Name?.Value, "Data", StringComparison.Ordinal));
            var dataPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(dataSheet.Id!);
            var chartPart = dataPart.DrawingsPart!.ChartParts.Single();
            var stylePart = Assert.Single(chartPart.GetPartsOfType<ChartStylePart>());
            var colorStylePart = Assert.Single(chartPart.GetPartsOfType<ChartColorStylePart>());
            using (var styleStream = stylePart.GetStream(FileMode.Open, FileAccess.Read)) {
                Assert.True(styleStream.Length > 0);
            }
            using (var colorStyleStream = colorStylePart.GetStream(FileMode.Open, FileAccess.Read)) {
                Assert.True(colorStyleStream.Length > 0);
            }
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void PerformanceReview_DashboardChartOverDataTableTableKeepsTheTableWorksheetData(bool includeCachedData) {
            var table = new DataTable();
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("Score", typeof(double));
            table.Columns.Add("Note", typeof(string));
            for (int i = 0; i < 40; i++) table.Rows.Add(i % 2 == 0 ? "North" : "South", i + 0.5, "n" + i);

            byte[] Save(bool addChart) {
                using var stream = new MemoryStream();
                using (var document = ExcelDocument.Create(stream, new ExcelCreateOptions {
                    PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
                })) {
                    var sheet = document.AddSheet("Data");
                    sheet.AddTable(table, new ExcelTableOptions { Name = "Scores", AutoFit = false }, row: 1, column: 1);
                    if (addChart) {
                        OfficeIMO.Excel.Reports.ExcelReportExtensions.AddDashboardChart(sheet, new OfficeIMO.Excel.Reports.ExcelDashboardChartOptions {
                            Range = "A1:B41", Row = 2, Column = 6, Title = "Scores", IncludeCachedData = includeCachedData
                        });
                    }
                }

                return stream.ToArray();
            }

            string SheetData(byte[] package) {
                using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
                using var reader = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml")!.Open());
                string xml = reader.ReadToEnd();
                int start = xml.IndexOf("<sheetData", StringComparison.Ordinal);
                return xml.Substring(start, xml.IndexOf("</sheetData>", StringComparison.Ordinal) - start);
            }

            byte[] withChart = Save(addChart: true);

            Assert.Equal(SheetData(Save(addChart: false)), SheetData(withChart));
            using var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(withChart), false);
            var chartPart = spreadsheet.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single();
            Assert.Equal(includeCachedData ? 40 : 0, chartPart.ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.NumericPoint>().Count());
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Theory]
        [InlineData("range")]
        [InlineData("tables")]
        public void PerformanceReview_TableMetadataReadsKeepTheDataTableWorksheetData(string read) {
            var table = new DataTable();
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Enabled", typeof(bool));
            table.Columns.Add("Created", typeof(DateTime));
            table.Columns.Add("Score", typeof(decimal));
            for (int i = 1; i <= 30; i++) table.Rows.Add(i, "n" + i, i % 2 == 0, new DateTime(2024, 1, 1).AddMinutes(i), i * 1.25m);

            string SheetData(bool readMetadata) {
                using var stream = new MemoryStream();
                using (var document = ExcelDocument.Create(stream, new ExcelCreateOptions {
                    PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
                })) {
                    var sheet = document.AddSheet("Data");
                    var created = sheet.AddTable(table, new ExcelTableOptions { Name = "Scores", AutoFit = false }, row: 1, column: 1);
                    if (readMetadata && read == "range") Assert.Equal("A1:E31", created.Range);
                    if (readMetadata && read == "tables") Assert.Equal("A1:E31", Assert.Single(sheet.GetTables()).Range);
                }

                using var archive = new ZipArchive(new MemoryStream(stream.ToArray()), ZipArchiveMode.Read);
                using var reader = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml")!.Open());
                string xml = reader.ReadToEnd();
                int start = xml.IndexOf("<sheetData", StringComparison.Ordinal);
                return xml.Substring(start, xml.IndexOf("</sheetData>", StringComparison.Ordinal) - start);
            }

            Assert.Equal(SheetData(readMetadata: false), SheetData(readMetadata: true));
        }

        [Fact]
        public void PerformanceReview_SheetViewEditsAfterFreezingADataTableKeepBothSettings() {
            var table = new DataTable();
            table.Columns.Add("Region", typeof(object));
            table.Columns.Add("Sales", typeof(object));
            table.Rows.Add("NA", 100);
            table.Rows.Add("EMEA", 200);

            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                var sheet = document.AddSheet("Data");
                var created = sheet.AddTable(table, new ExcelTableOptions { Name = "Sales", AutoFit = false }, row: 1, column: 1);
                Assert.Equal("A1:B3", created.Range);
                sheet.View.Freeze(1);
                sheet.View.ShowGridlines = false;
                document.Save(stream);
            }

            using var loaded = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
            var view = loaded.Sheets[0].View;
            Assert.False(view.ShowGridlines);
            Assert.Equal(1, view.FrozenRows);
            Assert.True(loaded.Sheets[0].TryGetCellText(3, 1, out string last));
            Assert.Equal("EMEA", last);
        }

        [Fact]
        public void PerformanceReview_EditsAfterReadingDataTableCellsReachTheSavedWorksheet() {
            var table = new DataTable();
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Name", typeof(object));
            for (int i = 1; i <= 30; i++) table.Rows.Add(i, "n" + i);

            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                var sheet = document.AddSheet("Data");
                sheet.AddTable(table, new ExcelTableOptions { Name = "Names", AutoFit = false }, row: 1, column: 1);
                Assert.True(sheet.TryGetCellText(3, 2, out string text));
                Assert.Equal("n2", text);
                sheet.Cell(1, 1).SetBold(true);
                sheet.Cell(5, 2).SetValue("edited");
                sheet.AutoFitColumnsFor(new[] { 1, 2 });
                document.Save(stream);
            }

            using var loaded = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
            var saved = loaded.Sheets[0];
            Assert.True(saved.TryGetCellText(5, 2, out string edited));
            Assert.Equal("edited", edited);
            Assert.True(saved.TryGetCellText(31, 2, out string last));
            Assert.Equal("n30", last);
            Assert.True(saved.Cell(1, 1).GetStyle().Bold);
        }

        [Fact]
        public void CellValueSnapshots_ColumnReadMatchesSingleCellReadsOnLoadedAndImportedSheets() {
            byte[] package;
            using (var stream = new MemoryStream()) {
                using (var document = ExcelDocument.Create(stream, new ExcelCreateOptions {
                    PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
                })) {
                    var sheet = document.AddSheet("Values");
                    sheet.Cell(1, 1).SetValue("Header");
                    sheet.Cell(2, 1).SetValue("Text");
                    sheet.Cell(3, 1).SetValue(42.5);
                    sheet.Cell(4, 1).SetValue(new DateTime(2024, 2, 29));
                    sheet.Cell(5, 1).SetValue(true);
                    sheet.Cell(6, 1).SetFormula("1+1");
                    sheet.Cell(8, 1).SetValue("Text");
                    sheet.Cell(9, 1).SetValue(-3);
                }

                package = stream.ToArray();
            }

            void AssertColumnMatches(ExcelSheet sheet, int column, int firstRow, int lastRow) {
                var snapshots = sheet.GetCellValueSnapshots(column, firstRow, lastRow);
                Assert.Equal(lastRow - firstRow + 1, snapshots.Count);
                for (int row = firstRow; row <= lastRow; row++) {
                    bool found = sheet.TryGetCellValueSnapshot(row, column, out var expected);
                    var actual = snapshots[row - firstRow];
                    Assert.Equal(found, actual != null);
                    if (!found) continue;
                    Assert.Equal(expected!.Kind, actual!.Kind);
                    Assert.Equal(expected.Text, actual.Text);
                    Assert.Equal(expected.RawValue, actual.RawValue);
                    Assert.Equal(expected.OpenXmlType, actual.OpenXmlType);
                    Assert.Equal(expected.DateTimeValue, actual.DateTimeValue);
                }
            }

            using (var loaded = ExcelDocument.Load(new MemoryStream(package))) {
                var sheet = loaded["Values"];
                AssertColumnMatches(sheet, 1, 1, 10);
                Assert.Equal(ExcelCellValueKind.DateTime, sheet.GetCellValueSnapshots(1, 4, 4)[0]!.Kind);
            }

            var table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Amount", typeof(double));
            table.Columns.Add("When", typeof(DateTime));
            for (int i = 0; i < 25; i++) table.Rows.Add("n" + (i % 4), i * 1.25, new DateTime(2020, 1, 1).AddDays(i));
            using var imported = ExcelDocument.Create(new MemoryStream());
            var importedSheet = imported.AddSheet("Imported");
            importedSheet.AddTable(table, new ExcelTableOptions { Name = "Imported", AutoFit = false }, row: 1, column: 1);
            for (int column = 1; column <= 3; column++) {
                AssertColumnMatches(importedSheet, column, 1, 27);
            }
        }

        [Fact]
        public void PerformanceReview_FileSnapshotKeepsAllViewsOnOneFileIdentity() {
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.Snapshot.{Guid.NewGuid():N}.xlsx");
            string replacementPath = path + ".replacement";
            byte[] original = Encoding.UTF8.GetBytes("original workbook snapshot");
            byte[] replacement = Encoding.UTF8.GetBytes("replacement workbook snapshot");
            try {
                File.WriteAllBytes(path, original);
                File.WriteAllBytes(replacementPath, replacement);
                using var snapshot = SharedReadOnlyFileSnapshot.Open(path);
                using Stream firstView = snapshot.CreateView();

                File.Delete(path);
                File.Move(replacementPath, path);

                using Stream secondView = snapshot.CreateView();
                Assert.Equal(original, ReadSnapshot(firstView));
                Assert.Equal(original, ReadSnapshot(secondView));
                Assert.Equal(replacement, File.ReadAllBytes(path));
            } finally {
                File.Delete(path);
                File.Delete(replacementPath);
            }
        }

        private static byte[] ReadSnapshot(Stream stream) {
            using var output = new MemoryStream();
            stream.CopyTo(output);
            return output.ToArray();
        }

        private static byte[] RewriteFirstWorksheetAsUtf16(byte[] packageBytes) {
            using var package = new MemoryStream();
            package.Write(packageBytes, 0, packageBytes.Length);
            package.Position = 0;

            using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
                const string worksheetPath = "xl/worksheets/sheet1.xml";
                var worksheetEntry = archive.GetEntry(worksheetPath)
                    ?? throw new InvalidOperationException("Worksheet entry was not found.");
                string worksheetXml;
                using (var reader = new StreamReader(worksheetEntry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true)) {
                    worksheetXml = reader.ReadToEnd();
                }

                worksheetEntry.Delete();
                var replacement = archive.CreateEntry(worksheetPath, CompressionLevel.Fastest);
                using var writer = new StreamWriter(replacement.Open(), new UnicodeEncoding(bigEndian: false, byteOrderMark: true));
                writer.Write(worksheetXml.Replace("utf-8", "utf-16").Replace("UTF-8", "utf-16"));
            }

            return package.ToArray();
        }

        private sealed class FormulaProjection {
            public FormulaProjection() {
            }

            public int Id { get; set; }
            public bool Active { get; set; }
            public string? Expression { get; set; }
        }

        private sealed class FastPackageProjection {
            internal FastPackageProjection(string name, int score) {
                Name = name;
                Score = score;
            }

            public string Name { get; }
            public int Score { get; }
        }
    }
}
