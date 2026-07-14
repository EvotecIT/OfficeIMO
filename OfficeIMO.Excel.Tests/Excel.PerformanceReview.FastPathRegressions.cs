using System;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
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
                PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose
            })) {
                var sheet = document.AddWorksheet("Formula");
                sheet.CellValue(2, 1, "Id");
                sheet.CellValue(2, 2, "Active");
                sheet.CellValue(2, 3, "Expression");
                sheet.CellFormula(3, 1, "1");
                sheet.CellFormula(3, 2, "TRUE");
                sheet.CellFormula(3, 3, "TRUE");
            }

            using var reader = ExcelDocumentReader.Open(memory.ToArray(), new ExcelReadOptions {
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
                PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose
            })) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Id");
                sheet.CellValue(1, 2, "Created");
                sheet.CellValue(1, 3, "Active");
                sheet.CellValue(1, 4, "Note");
                sheet.CellValue(2, 1, 42);
                sheet.CellValue(2, 2, expectedDate);
                sheet.CellValue(2, 3, true);
                sheet.CellValue(2, 4, "Complete");
                sheet.CellValue(3, 4, "Missing typed values");
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
                PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose
            })) {
                var rows = new[] {
                    new FastPackageProjection("Alpha", 10),
                    new FastPackageProjection("Beta", 20)
                };
                document.DefaultChartStylePreset = ExcelChartStylePreset.Default;
                var sheet = document.AddWorksheet("Data");
                sheet.InsertObjects(
                    rows,
                    ("Name", row => row.Name),
                    ("Score", row => row.Score));
                sheet.AddChart(
                    new ExcelChartData(
                        rows.Select(row => row.Name),
                        new[] { new ExcelChartSeries("Score", rows.Select(row => (double)row.Score)) }),
                    row: 1,
                    column: 5,
                    type: ExcelChartType.ColumnClustered,
                    title: "Scores");

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
