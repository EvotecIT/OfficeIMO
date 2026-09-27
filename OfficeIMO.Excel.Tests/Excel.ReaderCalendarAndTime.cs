using OfficeIMO.Excel;
using OfficeIMO.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private sealed class CalendarAndTimeRow {
            public DateTime Calendar { get; set; }
            public DateTime Duration { get; set; }
            public DateTime Time { get; set; }
            public double Numeric { get; set; }
            public string? Text { get; set; }
        }

        [Theory]
        [InlineData(ExcelDateSystem.NineteenHundred, false, false)]
        [InlineData(ExcelDateSystem.NineteenHundred, true, false)]
        [InlineData(ExcelDateSystem.NineteenFour, false, false)]
        [InlineData(ExcelDateSystem.NineteenFour, true, false)]
        [InlineData(ExcelDateSystem.NineteenHundred, false, true)]
        [InlineData(ExcelDateSystem.NineteenHundred, true, true)]
        [InlineData(ExcelDateSystem.NineteenFour, false, true)]
        [InlineData(ExcelDateSystem.NineteenFour, true, true)]
        public void Reader_CalendarAndTimeFormats_KeepDistinctSerialContracts(ExcelDateSystem system, bool converterFallback, bool unsorted) {
            string path = Path.Combine(_directoryWithFiles, $"CalendarAndTime{system}{converterFallback}{unsorted}.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                document.DateSystem = system;
                var sheet = document.AddWorksheet("Data");
                string[] headers = { "Calendar", "Duration", "Time", "Numeric", "Text" };
                for (int c = 0; c < headers.Length; c++) sheet.CellValue(1, c + 1, headers[c]);
                for (int row = 2; row <= 3; row++) {
                    sheet.CellValue(row, 1, row == 2 ? 1.5d : 59.5d);
                    sheet.CellValue(row, 2, 1.5d);
                    sheet.CellValue(row, 3, 0.5d);
                    sheet.CellValue(row, 4, 1.5d);
                    sheet.CellValue(row, 5, 1.5d);
                }
                sheet.ColumnStyleByHeader("Calendar").NumberFormat("yyyy-mm-dd hh:mm:ss");
                foreach (string header in new[] { "Duration", "Numeric", "Text" }) sheet.ColumnStyleByHeader(header).NumberFormat("[h]:mm");
                sheet.ColumnStyleByHeader("Time").NumberFormat("hh:mm:ss");
                document.Save();
            }
            if (unsorted) {
                using var package = SpreadsheetDocument.Open(path, true);
                var part = package.WorkbookPart!.WorksheetParts.Single();
                var data = part.Worksheet.GetFirstChild<SheetData>()!;
                var rows = data.Elements<Row>().ToArray();
                data.RemoveAllChildren<Row>();
                data.Append(rows[0], rows[2], rows[1]);
                part.Worksheet.Save();
            }
            var options = new ExcelReadOptions { SheetName = "Data" };
            if (converterFallback) options.CellValueConverter = _ => ExcelCellValue.NotHandled;
            DateTime first = ExcelDateSystemConverter.FromSerial(1.5d, system);
            DateTime second = ExcelDateSystemConverter.FromSerial(59.5d, system);
            DateTime duration = DateTime.FromOADate(1.5d);
            DateTime time = DateTime.FromOADate(0.5d);
            using (var reader = ExcelDocumentReader.Open(path, options)) {
                var sheet = reader.GetSheet("Data");
                var range = sheet.ReadRange("A2:E3");
                Assert.Equal(first, range[0, 0]); Assert.Equal(second, range[1, 0]);
                Assert.Equal(duration, range[0, 1]); Assert.Equal(time, range[0, 2]);
                var chunk = Assert.Single(sheet.ReadRangeStream("A2:E3", chunkRows: 2));
                Assert.Equal(duration, chunk.Rows[0][1]); Assert.Equal(time, chunk.Rows[0][2]);
                var rows = sheet.ReadObjectsStream<CalendarAndTimeRow>("A1:E3").ToArray();
                Assert.Equal(first, rows[0].Calendar); Assert.Equal(second, rows[1].Calendar);
                Assert.Equal(duration, rows[0].Duration); Assert.Equal(time, rows[0].Time);
                Assert.Equal(1.5d, rows[0].Numeric);
                Assert.Equal(duration.ToString(CultureInfo.InvariantCulture), rows[0].Text);
            }
            using var dataReader = ExcelDocument.OpenDataReader(path, options);
            Assert.True(dataReader.Read());
            Assert.Equal(1.5d, dataReader.GetDouble(dataReader.GetOrdinal("Numeric")));
            Assert.Equal(first, dataReader["Calendar"]); Assert.Equal(duration, dataReader["Duration"]); Assert.Equal(time, dataReader["Time"]);
            Assert.True(dataReader.Read()); Assert.Equal(second, dataReader["Calendar"]);
            Assert.False(dataReader.Read());
            using var mappedReader = ExcelDocument.OpenDataReader(path, options);
            Assert.All(mappedReader.RowsAs<CalendarAndTimeRow>(), row => Assert.Equal(1.5d, row.Numeric));
            using var parallelReader = ExcelDocument.OpenDataReader(path, options);
            Assert.All(parallelReader.RowsAsParallel<CalendarAndTimeRow>(new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 2 }), row => Assert.Equal(1.5d, row.Numeric));
        }

        [Theory]
        [InlineData(ExcelDateSystem.NineteenHundred)]
        [InlineData(ExcelDateSystem.NineteenFour)]
        public void Reader_CalendarAndTimeFormats_RespectMinuteTokenContext(ExcelDateSystem system) {
            string path = Path.Combine(_directoryWithFiles, $"CompactTime{system}.xlsx");
            string[] formats = { "hhmmss", "hhmm", "mmss", "hh mm", "mmhh", "hhmmm", "[mm]:ss", "[ss].00" };
            using (var document = ExcelDocument.Create(path)) {
                document.DateSystem = system;
                var sheet = document.AddWorksheet("Data");
                for (int c = 0; c < formats.Length; c++) {
                    sheet.CellValue(1, c + 1, "Field" + c);
                    sheet.CellValue(2, c + 1, 1.5d);
                    sheet.ColumnStyleByHeader("Field" + c).NumberFormat(formats[c]);
                }
                document.Save();
            }
            using var reader = ExcelDocumentReader.Open(path);
            var values = reader.GetSheet("Data").ReadRange("A2:H2");
            for (int c = 0; c < formats.Length; c++) {
                var expected = c is 4 or 5 ? ExcelDateSystemConverter.FromSerial(1.5d, system) : DateTime.FromOADate(1.5d);
                Assert.Equal(expected, values[0, c]);
            }
        }
    }
}
