using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Write;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Xlsb_LoadedWorkbook_RewritesEqualityListAutoFilterCriteria() {
            byte[] source;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(1, 2, "Owner");
                sheet.CellValue(2, 1, "Open");
                sheet.CellValue(2, 2, "Ada");
                sheet.CellValue(3, 1, "Closed");
                sheet.CellValue(3, 2, "Grace");
                sheet.AddAutoFilter("A1:B3", new Dictionary<uint, IEnumerable<string>> {
                    [0U] = new[] { "Open" }
                });
                source = document.ToBytes(ExcelFileFormat.Xlsb);
            }

            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source, writable: false));
            ExcelSheet loadedSheet = Assert.Single(loaded.Sheets);
            loadedSheet.AddAutoFilter("A1:B3", new Dictionary<uint, IEnumerable<string>> {
                [0U] = new[] { "Open", "Closed" },
                [1U] = new[] { "Ada" }
            });

            byte[] rewritten = loaded.ToBytes(ExcelFileFormat.Xlsb);
            Assert.Equal(
                ReadPackageEntry(source, "xl/workbook.bin"),
                ReadPackageEntry(rewritten, "xl/workbook.bin"));
            using (var archive = new ZipArchive(new MemoryStream(rewritten, writable: false), ZipArchiveMode.Read)) {
                using Stream worksheetStream = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("xl/worksheets/sheet1.bin")).Open();
                IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(worksheetStream);
                Assert.Single(records, record => record.Type == 161);
                Assert.Single(records, record => record.Type == 162);
                Assert.Equal(2, records.Count(record => record.Type == 163));
                Assert.Equal(3, records.Count(record => record.Type == 167));
            }

            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            AutoFilter filter = Assert.IsType<AutoFilter>(
                Assert.Single(reloaded.Sheets).WorksheetPart.Worksheet.GetFirstChild<AutoFilter>());
            Assert.Equal("A1:B3", filter.Reference?.Value);
            FilterColumn[] columns = filter.Elements<FilterColumn>().OrderBy(column => column.ColumnId?.Value).ToArray();
            Assert.Equal(2, columns.Length);
            Assert.Equal(new[] { "Open", "Closed" },
                columns[0].GetFirstChild<Filters>()!.Elements<Filter>().Select(value => value.Val?.Value));
            Assert.Equal(new[] { "Ada" },
                columns[1].GetFirstChild<Filters>()!.Elements<Filter>().Select(value => value.Val?.Value));
        }

        [Fact]
        public void Xlsb_LoadedWorkbook_ClearsEqualityListCriteriaWithoutRemovingFilterRange() {
            byte[] source;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(2, 1, "Open");
                sheet.AddAutoFilter("A1:A2", new Dictionary<uint, IEnumerable<string>> {
                    [0U] = new[] { "Open" }
                });
                source = document.ToBytes(ExcelFileFormat.Xlsb);
            }

            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source, writable: false));
            Assert.True(Assert.Single(loaded.Sheets).ClearAutoFilterColumn(0U));

            byte[] rewritten = loaded.ToBytes(ExcelFileFormat.Xlsb);
            using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(rewritten, writable: false));
            AutoFilter filter = Assert.IsType<AutoFilter>(
                Assert.Single(reloaded.Sheets).WorksheetPart.Worksheet.GetFirstChild<AutoFilter>());
            Assert.Equal("A1:A2", filter.Reference?.Value);
            Assert.Empty(filter.Elements<FilterColumn>());
        }

        [Fact]
        public void Xlsb_LoadedWorkbook_RewrittenAutoFilterOpensInDesktopExcelWhenAvailable() {
            if (!IsWindowsPlatform()) return;

            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Xlsb.AutoFilter.{Guid.NewGuid():N}.xlsb");
            try {
                byte[] source;
                using (ExcelDocument document = ExcelDocument.Create()) {
                    ExcelSheet sheet = document.AddWorksheet("Report");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(2, 1, "Open");
                    sheet.CellValue(3, 1, "Closed");
                    sheet.AddAutoFilter("A1:A3", new Dictionary<uint, IEnumerable<string>> {
                        [0U] = new[] { "Open" }
                    });
                    source = document.ToBytes(ExcelFileFormat.Xlsb);
                }

                using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source, writable: false))) {
                    Assert.Single(loaded.Sheets).AddAutoFilter(
                        "A1:A3",
                        new Dictionary<uint, IEnumerable<string>> {
                            [0U] = new[] { "Open", "Closed" }
                        });
                    File.WriteAllBytes(path, loaded.ToBytes(ExcelFileFormat.Xlsb));
                }

                AssertWorkbookOpensViaExcelComWhenAvailable(
                    path,
                    "The rewritten XLSB AutoFilter workbook failed to open in desktop Excel.");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Xlsb_LoadedWorkbook_PreservesRecognizedRecordsInsideUnsupportedAutoFilterAndBlocksMutation() {
            byte[] source;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(2, 1, "Open");
                sheet.AddAutoFilter("A1:A2", new Dictionary<uint, IEnumerable<string>> {
                    [0U] = new[] { "Open" }
                });
                source = document.ToBytes(ExcelFileFormat.Xlsb);
            }

            byte[] nonCanonical = InsertWorksheetRecordBeforeAutoFilterEnd(
                source,
                recordType: 477,
                payload: new byte[] { 0x00, 0x00 });
            using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(nonCanonical, writable: false))) {
                ExcelSheet sheet = Assert.Single(loaded.Sheets);
                sheet.CellValue(2, 1, "Edited");
                byte[] rewritten = loaded.ToBytes(ExcelFileFormat.Xlsb);
                IReadOnlyList<XlsbRecord> records = ReadWorksheetRecords(rewritten);
                int begin = records.ToList().FindIndex(record => record.Type == 161);
                int end = records.ToList().FindIndex(record => record.Type == 162);
                Assert.Contains(records.Skip(begin + 1).Take(end - begin - 1),
                    record => record.Type == 477 && record.Data.SequenceEqual(new byte[] { 0x00, 0x00 }));
            }

            using ExcelDocument changed = ExcelDocument.Load(new MemoryStream(nonCanonical, writable: false));
            Assert.Single(changed.Sheets).AddAutoFilter(
                "A1:A2",
                new Dictionary<uint, IEnumerable<string>> {
                    [0U] = new[] { "Open", "Closed" }
                });
            NotSupportedException exception = Assert.Throws<NotSupportedException>(
                () => changed.ToBytes(ExcelFileFormat.Xlsb));
            Assert.Contains("outside the supported equality-list subset", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Xlsb_LoadedWorkbook_RejectsForeignNamespaceAutoFilterAttributes() {
            byte[] source;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(2, 1, "Open");
                sheet.AddAutoFilter("A1:A2", new Dictionary<uint, IEnumerable<string>> {
                    [0U] = new[] { "Open" }
                });
                source = document.ToBytes(ExcelFileFormat.Xlsb);
            }

            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source, writable: false));
            ExcelSheet loadedSheet = Assert.Single(loaded.Sheets);
            loadedSheet.AddAutoFilter("A1:A2", new Dictionary<uint, IEnumerable<string>> {
                [0U] = new[] { "Open", "Closed" }
            });
            AutoFilter filter = Assert.IsType<AutoFilter>(
                loadedSheet.WorksheetPart.Worksheet.GetFirstChild<AutoFilter>());
            Filters values = Assert.IsType<Filters>(Assert.Single(filter.Elements<FilterColumn>()).GetFirstChild<Filters>());
            values.SetAttribute(new OpenXmlAttribute("ext", "blank", "urn:officeimo:test", "1"));

            NotSupportedException exception = Assert.Throws<NotSupportedException>(
                () => loaded.ToBytes(ExcelFileFormat.Xlsb));
            Assert.Contains("attribute 'blank'", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Xlsb_NewWorkbook_AllowsRelationshipQualifiedHyperlinkIdButRejectsUnqualifiedId() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Report");
            sheet.CellValue(1, 1, "OfficeIMO");
            sheet.SetHyperlink(
                1,
                1,
                "https://evotec.xyz/",
                display: "OfficeIMO",
                style: false);

            Assert.NotEmpty(document.ToBytes(ExcelFileFormat.Xlsb));
            Hyperlink hyperlink = Assert.Single(
                sheet.WorksheetPart.Worksheet.GetFirstChild<Hyperlinks>()!.Elements<Hyperlink>());
            Assert.False(string.IsNullOrEmpty(hyperlink.Id?.Value));
            hyperlink.SetAttribute(new OpenXmlAttribute(string.Empty, "id", string.Empty, "rIdShadow"));

            NotSupportedException exception = Assert.Throws<NotSupportedException>(
                () => document.ToBytes(ExcelFileFormat.Xlsb));
            Assert.Contains("attribute 'id'", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Xlsb_NewWorkbook_RejectsUnqualifiedPageSetupId() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Report");
            sheet.CellValue(1, 1, "Status");
            sheet.SetPageSetup(fitToWidth: 1U, fitToHeight: 0U);
            PageSetup pageSetup = Assert.IsType<PageSetup>(
                sheet.WorksheetPart.Worksheet.GetFirstChild<PageSetup>());
            pageSetup.SetAttribute(new OpenXmlAttribute(string.Empty, "id", string.Empty, "rIdShadow"));

            NotSupportedException exception = Assert.Throws<NotSupportedException>(
                () => document.ToBytes(ExcelFileFormat.Xlsb));
            Assert.Contains("attribute 'id'", exception.Message, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Xlsb_LoadedWorkbook_FailsClosedWhenFilterDatabaseNameWouldNeedChanging(bool remove) {
            byte[] source;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(2, 1, "Open");
                sheet.AddAutoFilter("A1:A2");
                source = document.ToBytes(ExcelFileFormat.Xlsb);
            }

            using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source, writable: false));
            ExcelSheet loadedSheet = Assert.Single(loaded.Sheets);
            if (remove) {
                loadedSheet.AutoFilterClear();
            } else {
                loadedSheet.AutoFilterAdd("A1:B2");
            }

            NotSupportedException exception = Assert.Throws<NotSupportedException>(
                () => loaded.ToBytes(ExcelFileFormat.Xlsb));
            Assert.Contains("_FilterDatabase", exception.Message, StringComparison.Ordinal);
        }

        private static byte[] ReadPackageEntry(byte[] package, string entryName) {
            using var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
            using Stream stream = Assert.IsType<ZipArchiveEntry>(archive.GetEntry(entryName)).Open();
            using var output = new MemoryStream();
            stream.CopyTo(output);
            return output.ToArray();
        }

        private static IReadOnlyList<XlsbRecord> ReadWorksheetRecords(byte[] package) {
            byte[] worksheet = ReadPackageEntry(package, "xl/worksheets/sheet1.bin");
            using var stream = new MemoryStream(worksheet, writable: false);
            return XlsbRecordReader.ReadAll(stream);
        }

        private static byte[] InsertWorksheetRecordBeforeAutoFilterEnd(
            byte[] package,
            int recordType,
            byte[] payload) {
            IReadOnlyList<XlsbRecord> records = ReadWorksheetRecords(package);
            using var worksheet = new MemoryStream();
            bool inserted = false;
            foreach (XlsbRecord record in records) {
                if (!inserted && record.Type == 162) {
                    XlsbRecordWriter.Write(worksheet, recordType, payload);
                    inserted = true;
                }
                XlsbRecordWriter.Write(worksheet, record.Type, record.Data);
            }
            Assert.True(inserted);
            return XlsbNativePackageWriter.RewritePackage(
                package,
                new Dictionary<string, byte[]> {
                    ["xl/worksheets/sheet1.bin"] = worksheet.ToArray()
                },
                maxPartBytes: 128 * 1024 * 1024,
                maxPackageBytes: 512L * 1024 * 1024);
        }
    }
}
