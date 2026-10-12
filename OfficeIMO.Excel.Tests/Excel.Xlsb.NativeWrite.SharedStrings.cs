using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.Data;
using System.IO.Compression;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(null, false)]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(null, true)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public async Task Xlsb_StringStorage_SelectsNativeRecordsAndPreservesTabularValues(bool? sharedStrings, bool standardWriter) {
            string?[] texts = { "Alpha", " Alpha & <β>\r ", "", null, "Alpha", "alpha" };
            var table = new DataTable("Data");
            table.Columns.Add("Text", typeof(string));
            table.Columns.Add("Id", typeof(int));
            table.Columns.Add("Enabled", typeof(bool));
            for (int index = 0; index < texts.Length; index++) {
                table.Rows.Add(texts[index] == null ? DBNull.Value : texts[index], index + 1, index % 2 == 0);
            }
            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            using ExcelDocument document = ExcelDocument.Create();
            document.InsertDataSet(dataSet, createTables: false, includeAutoFilter: false);
            var options = new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings, DisableFastPackageWriter = standardWriter };
            using var destination = new MemoryStream();

            await document.SaveAsync(destination, ExcelFileFormat.Xlsb, options);

            Assert.Equal(sharedStrings, options.XlsbUseSharedStrings);
            Assert.Equal(!standardWriter, document.LastSaveDiagnostics.UsedFastPackageWriter);
            byte[] package = destination.ToArray();
            AssertXlsbStringStorage(package, sharedStrings == true, 8,
                new[] { "Text", "Id", "Enabled", "Alpha", " Alpha & <β>\r ", "", "alpha" });
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package);
            Assert.Equal(new[] { "Text", "Id", "Enabled" }, Enumerable.Range(0, reader.FieldCount).Select(reader.GetName));
            for (int index = 0; index < texts.Length; index++) {
                Assert.True(reader.Read());
                if (texts[index] == null) {
                    Assert.True(reader.IsDBNull(0));
                } else {
                    Assert.False(reader.IsDBNull(0));
                    Assert.Equal(texts[index], reader.GetString(0));
                }
                Assert.Equal(index + 1, reader.GetInt32(1));
                Assert.Equal(index % 2 == 0, reader.GetBoolean(2));
            }
            Assert.False(reader.Read());
            Assert.False(reader.NextResult());
        }

        [Fact]
        public void Xlsb_SharedStrings_DeduplicateAcrossSheetsWithoutChangingStylesDatesOrFormulaCaches() {
            DateTime date = new DateTime(2026, 10, 8, 6, 0, 0);
            using ExcelDocument document = ExcelDocument.Create();
            document.DateSystem = ExcelDateSystem.NineteenFour;
            ExcelSheet first = document.AddWorksheet("First");
            first.CellValue(1, 1, "Alpha");
            first.CellBold(1, 1);
            first.CellValue(2, 1, date);
            first.CellValue(3, 1, 12.5m);
            first.FormatCell(3, 1, "0.0000");
            first.CellFormula(4, 1, "\"Alpha\"");
            Cell formula = first.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "A4");
            formula.DataType = CellValues.String;
            formula.CellValue = new CellValue("Alpha");
            ExcelSheet second = document.AddWorksheet("Second");
            second.CellValue(1, 1, "Alpha");
            second.CellValue(2, 1, "alpha");

            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb, new ExcelSaveOptions { XlsbUseSharedStrings = true });

            AssertXlsbStringStorage(package, true, 3, new[] { "Alpha", "alpha" });
            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            Assert.Equal(ExcelDateSystem.NineteenFour, reloaded.DateSystem);
            Assert.True(reloaded.Sheets[0].GetCellStyle(1, 1).Bold);
            Assert.Equal(date, AssertCellValue(reloaded.Sheets[0], 2, 1).DateTimeValue);
            Assert.Equal("0.0000", reloaded.Sheets[0].GetCellStyle(3, 1).NumberFormatCode);
            Assert.Equal("\"Alpha\"", reloaded.Sheets[0].CellAt(4, 1).GetValue().Formula);
            Assert.True(reloaded.Sheets[0].TryGetCellText(4, 1, out string? cached));
            Assert.Equal("Alpha", cached);
            Assert.True(reloaded.Sheets[1].TryGetCellText(2, 1, out string? lowercase));
            Assert.Equal("alpha", lowercase);
        }

        [Fact]
        public async Task Xlsb_SharedStrings_FileSaveAdoptsStateAndDefaultRewritePreservesTheTable() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
            try {
                using ExcelDocument document = ExcelDocument.Create();
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.CellValue(2, 1, "Alpha");
                await document.SaveAsync(path, new ExcelSaveOptions { XlsbUseSharedStrings = true });
                byte[] original = File.ReadAllBytes(path);
                byte[] originalTable = ReadXlsbPackagePart(original, "xl/sharedStrings.bin");
                Assert.Equal(ExcelFileFormat.Xlsb, document.SourceFormat);
                using var unchanged = new MemoryStream();
                document.Save(unchanged, ExcelFileFormat.Xlsb);
                Assert.Equal(original, unchanged.ToArray());

                sheet.CellValue(2, 1, "Beta");
                document.Save();

                byte[] rewritten = File.ReadAllBytes(path);
                Assert.Equal(originalTable, ReadXlsbPackagePart(rewritten, "xl/sharedStrings.bin"));
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(rewritten);
                Assert.True(reader.Read());
                Assert.Equal("Beta", reader.GetString(0));
                Assert.False(reader.Read());
            } finally {
                TryDelete(path);
            }
        }

        [Theory]
        [InlineData(true, false)]
        [InlineData(false, false)]
        [InlineData(true, true)]
        [InlineData(false, true)]
        public async Task Xlsb_StringStorage_RejectsImportedOverridesBeforeMutatingStreamOrFile(bool sharedStrings, bool asynchronous) {
            using ExcelDocument source = ExcelDocument.Create();
            source.AddWorksheet("Data").CellValue(1, 1, "Alpha");
            byte[] package = source.ToBytes(ExcelFileFormat.Xlsb, new ExcelSaveOptions { XlsbUseSharedStrings = true });
            using ExcelDocument imported = ExcelDocument.Load(new MemoryStream(package, writable: false));
            if (!sharedStrings) imported.Sheets[0].CellValue(1, 1, "Edited");
            var options = new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings };
            byte[] sentinel = { 1, 2, 3, 4 };
            using var destination = new MemoryStream();
            destination.Write(sentinel, 0, sentinel.Length);
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
            File.WriteAllBytes(path, sentinel);
            try {
                NotSupportedException streamError;
                NotSupportedException fileError;
                if (asynchronous) {
                    streamError = await Assert.ThrowsAsync<NotSupportedException>(() => imported.SaveAsync(destination, ExcelFileFormat.Xlsb, options));
                    fileError = await Assert.ThrowsAsync<NotSupportedException>(() => imported.SaveAsync(path, options));
                } else {
                    streamError = Assert.Throws<NotSupportedException>(() => imported.Save(destination, ExcelFileFormat.Xlsb, options));
                    fileError = Assert.Throws<NotSupportedException>(() => imported.Save(path, options));
                }
                Assert.Contains("imported XLSB", streamError.Message, StringComparison.Ordinal);
                Assert.Contains("imported XLSB", fileError.Message, StringComparison.Ordinal);
                Assert.Equal(sentinel, destination.ToArray());
                Assert.Equal(sentinel.Length, destination.Position);
                Assert.Equal(sentinel, File.ReadAllBytes(path));
            } finally {
                TryDelete(path);
            }
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task Xlsb_StringStorage_RejectsOtherOutputFormatsAndEncryptedOutput(bool sharedStrings) {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Data").CellValue(1, 1, "Alpha");
            var options = new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings };
            using var destination = new MemoryStream();
            destination.WriteByte(42);
            Assert.Throws<NotSupportedException>(() => document.Save(destination, ExcelFileFormat.Xlsx, options));
            await Assert.ThrowsAsync<NotSupportedException>(() => document.SaveAsync(destination, ExcelFileFormat.Xls, options));
            Assert.Throws<NotSupportedException>(() => document.SaveEncrypted(destination, "test password", options));
            Assert.Equal(new byte[] { 42 }, destination.ToArray());
            Assert.Equal(1, destination.Position);
        }

        [Fact]
        public void Xlsb_SharedStrings_PreserveMaximumLengthUtf16AcrossBufferBoundaries() {
            string text = new string('A', 2_047) + char.ConvertFromUtf32(0x1F680) + new string('Z', 32_767 - 2_049);
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            sheet.CellValue(2, 1, text);
            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb, new ExcelSaveOptions { XlsbUseSharedStrings = true });
            AssertXlsbStringStorage(package, true, 2, new[] { "Value", text });
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package);
            Assert.True(reader.Read());
            Assert.Equal(text, reader.GetString(0));
            Assert.False(reader.Read());
        }

        [Fact]
        public void Xlsb_SharedStrings_AllowAnEmptyTableForNumericOnlyWorkbooks() {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Numbers").CellValue(1, 1, 42);
            byte[] package = document.ToBytes(ExcelFileFormat.Xlsb, new ExcelSaveOptions { XlsbUseSharedStrings = true });
            AssertXlsbStringStorage(package, true, 0, Array.Empty<string>());
            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
            ExcelCellValueSnapshot value = AssertCellValue(reloaded.Sheets[0], 1, 1);
            Assert.Equal(ExcelCellValueKind.Number, value.Kind);
            Assert.Equal("42", value.RawValue);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Xlsb_SharedStrings_OpenInDesktopExcelWhenAvailable(bool standardWriter) {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
            try {
                using ExcelDocument document = ExcelDocument.Create();
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Text");
                sheet.CellValue(2, 1, "Alpha");
                sheet.CellValue(3, 1, "Alpha");
                sheet.CellValue(4, 1, "alpha");
                if (standardWriter) {
                    sheet.CellBold(1, 1);
                    sheet.CellFormula(5, 1, "\"Alpha\"");
                }
                document.Save(path, new ExcelSaveOptions { XlsbUseSharedStrings = true, DisableFastPackageWriter = standardWriter });
                AssertWorkbookOpensViaExcelComWhenAvailable(path,
                    "The XLSB shared-string table must be readable by desktop Excel.",
                    new Dictionary<string, string> { ["A1"] = "Text", ["A2"] = "Alpha", ["A3"] = "Alpha", ["A4"] = "alpha" });
            } finally {
                TryDelete(path);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Xlsb_StringStorage_ConversionCarriesTheSelectedSaveOption(bool sharedStrings) {
            string sourcePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string destinationPath = Path.ChangeExtension(sourcePath, ".xlsb");
            try {
                using (ExcelDocument source = ExcelDocument.Create()) {
                    ExcelSheet sheet = source.AddWorksheet("Data");
                    sheet.CellValue(1, 1, "Text");
                    sheet.CellValue(2, 1, "Alpha");
                    sheet.CellValue(3, 1, "Alpha");
                    source.Save(sourcePath);
                }
                ExcelDocument.Convert(sourcePath, destinationPath, new ExcelDocumentConversionOptions {
                    SaveOptions = new ExcelSaveOptions { XlsbUseSharedStrings = sharedStrings }
                });
                byte[] package = File.ReadAllBytes(destinationPath);
                AssertXlsbStringStorage(package, sharedStrings, 3, new[] { "Text", "Alpha" });
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(package);
                Assert.True(reader.Read());
                Assert.Equal("Alpha", reader.GetString(0));
                Assert.True(reader.Read());
                Assert.Equal("Alpha", reader.GetString(0));
                Assert.False(reader.Read());
            } finally {
                TryDelete(sourcePath);
                TryDelete(destinationPath);
            }
        }

        private static void AssertXlsbStringStorage(byte[] package, bool sharedStrings, int referenceCount, string[] values) {
            using var input = new MemoryStream(package, writable: false);
            using var archive = new ZipArchive(input, ZipArchiveMode.Read);
            ZipArchiveEntry? part = archive.GetEntry("xl/sharedStrings.bin");
            var cells = new List<XlsbRecord>();
            foreach (ZipArchiveEntry worksheet in archive.Entries.Where(entry => entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal) && entry.FullName.EndsWith(".bin", StringComparison.Ordinal))) {
                using Stream stream = worksheet.Open();
                cells.AddRange(XlsbRecordReader.ReadAll(stream).Where(record => record.Type == 6 || record.Type == 7));
            }
            Assert.Equal(referenceCount, cells.Count);
            Assert.All(cells, record => Assert.Equal(sharedStrings ? 7 : 6, record.Type));
            XNamespace relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
            using Stream relationships = archive.GetEntry("xl/_rels/workbook.bin.rels")!.Open();
            XElement? relationship = XDocument.Load(relationships).Root!.Elements(relationshipNamespace + "Relationship")
                .SingleOrDefault(element => ((string?)element.Attribute("Type"))?.EndsWith("/sharedStrings", StringComparison.Ordinal) == true);
            using Stream types = archive.GetEntry("[Content_Types].xml")!.Open();
            XElement? contentType = XDocument.Load(types).Root!.Elements().SingleOrDefault(element => (string?)element.Attribute("PartName") == "/xl/sharedStrings.bin");
            if (!sharedStrings) {
                Assert.Null(part);
                Assert.Null(relationship);
                Assert.Null(contentType);
                return;
            }
            Assert.NotNull(part);
            Assert.Equal("sharedStrings.bin", (string?)relationship?.Attribute("Target"));
            Assert.Equal("application/vnd.ms-excel.sharedStrings", (string?)contentType?.Attribute("ContentType"));
            using Stream sharedStringStream = part!.Open();
            IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(sharedStringStream);
            Assert.Equal(159, records[0].Type);
            var counts = new XlsbBinaryCursor(records[0].Data);
            Assert.Equal((uint)referenceCount, counts.ReadUInt32());
            Assert.Equal((uint)values.Length, counts.ReadUInt32());
            Assert.Equal(0, counts.Remaining);
            Assert.Equal(values.Length + 2, records.Count);
            Assert.Equal(160, records[records.Count - 1].Type);
            Assert.Empty(records[records.Count - 1].Data);
            for (int index = 0; index < values.Length; index++) {
                Assert.Equal(19, records[index + 1].Type);
                var value = new XlsbBinaryCursor(records[index + 1].Data);
                Assert.Equal(0, value.ReadByte());
                Assert.Equal(values[index], value.ReadWideString(32_767));
                Assert.Equal(0, value.Remaining);
            }
            Assert.All(cells, cell => {
                Assert.Equal(12, cell.Data.Length);
                var cursor = new XlsbBinaryCursor(cell.Data);
                cursor.Skip(8);
                Assert.InRange(cursor.ReadUInt32(), 0U, checked((uint)values.Length - 1));
            });
        }

        private static byte[] ReadXlsbPackagePart(byte[] package, string name) {
            using var input = new MemoryStream(package, writable: false);
            using var archive = new ZipArchive(input, ZipArchiveMode.Read);
            using Stream part = archive.GetEntry(name)!.Open();
            using var result = new MemoryStream();
            part.CopyTo(result);
            return result.ToArray();
        }
    }
}
