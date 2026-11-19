using System.Data;
using System.Globalization;
using System.Text;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests {
    public partial class Excel {
        public sealed class IsoDateCellRow {
            public DateTime Time { get; set; }
            public DateTime OffsetTime { get; set; }
            public DateTime Date { get; set; }
        }

        [Theory]
        [InlineData(false, false, "en-US")]
        [InlineData(true, false, "en-US")]
        [InlineData(false, true, "en-US")]
        [InlineData(true, true, "th-TH")]
        [InlineData(false, false, "th-TH")]
        public void IsoDateCellsReadTimeOfDayWithoutTodaysDateAndIgnoreCalendarCulture(bool date1904, bool utf16Worksheet, string cultureName) {
            string path = CreateIsoDateCellWorkbook(date1904, utf16Worksheet);
            CultureInfo originalCulture = CultureInfo.CurrentCulture;
            var culture = CultureInfo.GetCultureInfo(cultureName);
            try {
                CultureInfo.CurrentCulture = culture;
                // A time-only value reads as the same time-only cell stored as a serial with a time format does.
                DateTime expectedTime = DateTime.FromOADate(0.4375);
                var expectedDate = new DateTime(2024, 3, 5);
                var options = new ExcelReaderOptions { Culture = culture };

                using (var reader = ExcelDocumentReader.Open(path, options)) {
                    ExcelSheetReader sheet = reader.GetSheet("Data");
                    object?[,] range = sheet.ReadRange("A2:D2");
                    Assert.Equal(expectedTime, Assert.IsType<DateTime>(range[0, 0]));
                    Assert.Equal(expectedTime, Assert.IsType<DateTime>(range[0, 1]));
                    Assert.Equal(expectedDate, Assert.IsType<DateTime>(range[0, 2]));
                    Assert.Equal(expectedTime, Assert.IsType<DateTime>(range[0, 3]));

                    IsoDateCellRow row = sheet.ReadObjects<IsoDateCellRow>("A1:C2").Single();
                    Assert.Equal(expectedTime, row.Time);
                    Assert.Equal(expectedTime, row.OffsetTime);
                    Assert.Equal(expectedDate, row.Date);

                    IsoDateCellRow streamed = sheet.ReadObjectsStream<IsoDateCellRow>("A1:C2").Single();
                    Assert.Equal(expectedTime, streamed.Time);
                    Assert.Equal(expectedTime, streamed.OffsetTime);
                    Assert.Equal(expectedDate, streamed.Date);

                    using IDataReader rangeReader = sheet.ReadRangeAsDataReader("A1:D2");
                    AssertIsoDateCellRecord(rangeReader, expectedTime, expectedDate);
                }

                using (ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path, options)) {
                    AssertIsoDateCellRecord(reader, expectedTime, expectedDate);
                }

                using (var document = ExcelDocument.Load(path)) {
                    ExcelSheet sheet = document.Sheets.Single(s => s.Name == "Data");
                    Assert.Equal("10:30:00", sheet.Cell(2, 1).GetFormattedText(CultureInfo.InvariantCulture));
                    Assert.Equal("10:30:00", sheet.Cell(2, 2).GetFormattedText(CultureInfo.InvariantCulture));
                    Assert.Equal(sheet.Cell(2, 4).GetFormattedText(CultureInfo.InvariantCulture),
                        sheet.Cell(2, 2).GetFormattedText(CultureInfo.InvariantCulture));
                }

                using (ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path, options)) {
                    IsoDateCellRow mapped = reader.RowsAs<IsoDateCellRow>().Single();
                    Assert.Equal(expectedTime, mapped.Time);
                    Assert.Equal(expectedTime, mapped.OffsetTime);
                    Assert.Equal(expectedDate, mapped.Date);
                }
            } finally {
                CultureInfo.CurrentCulture = originalCulture;
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(".xls", false)]
        [InlineData(".xls", true)]
        [InlineData(".xlsb", false)]
        [InlineData(".xlsb", true)]
        public void IsoTimeOnlyCellsConvertToTimeOfDaySerialsInBinaryFormats(string extension, bool date1904) {
            string path = CreateIsoDateCellWorkbook(date1904, utf16Worksheet: false);
            string converted = Path.ChangeExtension(path, extension);
            try {
                using (var document = ExcelDocument.Load(path)) {
                    document.Save(converted, new ExcelSaveOptions { Format = extension == ".xls" ? ExcelFileFormat.Xls : ExcelFileFormat.Xlsb });
                }

                using var convertedDocument = ExcelDocument.Load(converted);
                ExcelSheet sheet = convertedDocument.Sheets.Single(s => s.Name == "Data");
                Assert.Equal(0.4375D, sheet.Cell(2, 1).GetValue<double>());
                Assert.Equal(0.4375D, sheet.Cell(2, 2).GetValue<double>());
                Assert.Equal(ExcelDateSystemConverter.ToSerial(new DateTime(2024, 3, 5), convertedDocument.DateSystem),
                    sheet.Cell(2, 3).GetValue<double>());
            } finally {
                File.Delete(path);
                File.Delete(converted);
            }
        }

        private static void AssertIsoDateCellRecord(IDataReader reader, DateTime expectedTime, DateTime expectedDate) {
            Assert.True(reader.Read());
            Assert.Equal(expectedTime, reader.GetDateTime(0));
            Assert.Equal(expectedTime, Assert.IsType<DateTime>(reader.GetValue(0)));
            Assert.Equal(expectedTime, reader.GetDateTime(1));
            Assert.Equal(expectedDate, reader.GetDateTime(2));
            Assert.Equal(expectedTime, reader.GetDateTime(3));
        }

        private static string CreateIsoDateCellWorkbook(bool date1904, bool utf16Worksheet) {
            string path = CreateCompactFastPathWorkbook();
            try {
                string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<dimension ref=\"A1:D2\"/><sheetData><row r=\"1\">"
                    + "<c r=\"A1\" t=\"inlineStr\"><is><t>Time</t></is></c>"
                    + "<c r=\"B1\" t=\"inlineStr\"><is><t>OffsetTime</t></is></c>"
                    + "<c r=\"C1\" t=\"inlineStr\"><is><t>Date</t></is></c>"
                    + "<c r=\"D1\" t=\"inlineStr\"><is><t>Numeric</t></is></c></row>"
                    + "<row r=\"2\"><c r=\"A2\" s=\"1\" t=\"d\"><v>10:30:00</v></c>"
                    + "<c r=\"B2\" s=\"1\" t=\"d\"><v>10:30:00-03:00</v></c>"
                    + "<c r=\"C2\" s=\"2\" t=\"d\"><v>2024-03-05</v></c>"
                    + "<c r=\"D2\" s=\"1\"><v>0.4375</v></c></row></sheetData></worksheet>";
                byte[] bytes = utf16Worksheet
                    ? Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("<?xml version=\"1.0\" encoding=\"utf-16\"?>" + xml)).ToArray()
                    : Encoding.UTF8.GetBytes(xml);
                ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", bytes);
                ReplaceZipEntry(path, "xl/styles.xml", Encoding.UTF8.GetBytes(
                    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders>"
                    + "<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>"
                    + "<cellXfs count=\"3\"><xf numFmtId=\"0\"/><xf numFmtId=\"21\" applyNumberFormat=\"1\"/><xf numFmtId=\"14\" applyNumberFormat=\"1\"/></cellXfs></styleSheet>"));
                if (date1904) {
                    string workbook;
                    using (var archive = System.IO.Compression.ZipFile.OpenRead(path))
                    using (var streamReader = new StreamReader(archive.GetEntry("xl/workbook.xml")!.Open())) {
                        workbook = streamReader.ReadToEnd();
                    }
                    workbook = workbook.Contains("<workbookPr")
                        ? workbook.Replace("<workbookPr", "<workbookPr date1904=\"1\"")
                        : workbook.Replace("<sheets", "<workbookPr date1904=\"1\"/><sheets");
                    ReplaceZipEntry(path, "xl/workbook.xml", Encoding.UTF8.GetBytes(workbook));
                }
                return path;
            } catch {
                File.Delete(path);
                throw;
            }
        }
    }
}
