using System.Data;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public sealed class EndOfDayTimeRow {
            public TimeSpan Duration { get; set; }
            public DateTime When { get; set; }
            public TimeSpan Elapsed { get; set; }
            public TimeSpan Negative { get; set; }
#if NET6_0_OR_GREATER
            public TimeOnly Time { get; set; }
#endif
        }

        /// <summary>The serial Excel stores for a duration: whole milliseconds as a fraction of a day.</summary>
        internal static double MillisecondSerial(TimeSpan value) => value.Ticks / TimeSpan.TicksPerMillisecond / 86_400_000d;

        private static readonly long LastTickOfDay = TimeSpan.TicksPerDay - 1;
        private static readonly DateTime TimeSerialBase = new DateTime(1899, 12, 30);
        private static readonly TimeSpan LastMillisecondOfDay = new TimeSpan(TimeSpan.TicksPerDay - TimeSpan.TicksPerMillisecond);

        private static EndOfDayTimeRow CreateEndOfDayTimeRow() => new EndOfDayTimeRow {
            Duration = new TimeSpan(LastTickOfDay),
            When = new DateTime(2024, 3, 5).AddTicks(LastTickOfDay),
            Elapsed = new TimeSpan(TimeSpan.TicksPerDay + LastTickOfDay),
            Negative = new TimeSpan(-LastTickOfDay),
#if NET6_0_OR_GREATER
            Time = TimeOnly.MaxValue,
#endif
        };

        // Values keep whole milliseconds, truncated toward zero, and never move into the next day. Time and
        // elapsed-time cells read as their duration after December 30, 1899 in both date systems.
        private static IReadOnlyList<(string Name, DateTime Expected)> ExpectedEndOfDayTimes() {
            var expected = new List<(string, DateTime)> {
                ("Duration", TimeSerialBase + LastMillisecondOfDay),
                ("When", new DateTime(2024, 3, 5) + LastMillisecondOfDay),
                ("Elapsed", TimeSerialBase.AddDays(1) + LastMillisecondOfDay),
                ("Negative", TimeSerialBase - LastMillisecondOfDay),
            };
#if NET6_0_OR_GREATER
            expected.Add(("Time", TimeSerialBase + LastMillisecondOfDay));
#endif
            return expected;
        }

        private static DataTable CreateEndOfDayTimeTable() {
            EndOfDayTimeRow row = CreateEndOfDayTimeRow();
            var table = new DataTable("Times");
            table.Columns.Add("Duration", typeof(TimeSpan));
            table.Columns.Add("When", typeof(DateTime));
            table.Columns.Add("Elapsed", typeof(TimeSpan));
            table.Columns.Add("Negative", typeof(TimeSpan));
#if NET6_0_OR_GREATER
            table.Columns.Add("Time", typeof(TimeOnly));
            table.Rows.Add(row.Duration, row.When, row.Elapsed, row.Negative, row.Time);
#else
            table.Rows.Add(row.Duration, row.When, row.Elapsed, row.Negative);
#endif
            return table;
        }

        public static IEnumerable<object[]> EndOfDayTimeWriteRoutes() {
            foreach (bool date1904 in new[] { false, true }) {
                foreach (ExcelFileFormat format in new[] { ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb, ExcelFileFormat.Xls }) {
                    yield return new object[] { "DataTable", format, date1904 };
                    yield return new object[] { "DataReader", format, date1904 };
                    yield return new object[] { "Objects", format, date1904 };
                }
                yield return new object[] { "CellValues", ExcelFileFormat.Xlsx, date1904 };
                yield return new object[] { "WriteData", ExcelFileFormat.Xlsx, date1904 };
            }
        }

        [Theory]
        [MemberData(nameof(EndOfDayTimeWriteRoutes))]
        public void EndOfDayTimesStayOnTheirDayThroughWriteAndRead(string route, ExcelFileFormat format, bool date1904) {
            ExcelDateSystem dateSystem = date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
            string path = Path.Combine(_directoryWithFiles, $"EndOfDay-{route}-{dateSystem}.{format.ToString().ToLowerInvariant()}");
            WriteEndOfDayTimes(route, format, dateSystem, path);
            IReadOnlyList<(string Name, DateTime Expected)> expected = ExpectedEndOfDayTimes();

            using (ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path)) {
                Assert.True(reader.Read());
                foreach ((string name, DateTime value) in expected) {
                    Assert.Equal(value, reader.GetDateTime(reader.GetOrdinal(name)));
                }
            }

            if (format != ExcelFileFormat.Xlsx) return;

            using (var document = ExcelDocumentReader.Open(path)) {
                object?[,] range = document.GetSheet("Times").ReadRange("A2:" + (char)('A' + expected.Count - 1) + "2");
                for (int column = 0; column < expected.Count; column++) {
                    Assert.Equal(expected[column].Expected, Assert.IsType<DateTime>(range[0, column]));
                }
            }

            using (var spreadsheet = SpreadsheetDocument.Open(path, false)) {
                Cell duration = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>()
                    .First(cell => cell.CellReference?.Value == "A2");
                Assert.Equal(86_399_999d / 86_400_000d, double.Parse(duration.CellValue!.Text, CultureInfo.InvariantCulture));
            }
        }

        private static void WriteEndOfDayTimes(string route, ExcelFileFormat format, ExcelDateSystem dateSystem, string path) {
            if (route == "CellValues" || route == "WriteData") {
                using var document = ExcelDocument.Create(path);
                document.DateSystem = dateSystem;
                ExcelSheet sheet = document.AddSheet("Times");
                DataTable table = CreateEndOfDayTimeTable();
                if (route == "WriteData") {
                    sheet.WriteData(table, includeHeaders: true);
                } else {
                    for (int column = 0; column < table.Columns.Count; column++) {
                        sheet.Cell(1, column + 1).SetValue(table.Columns[column].ColumnName);
                        sheet.Cell(2, column + 1).SetValue(table.Rows[0][column]);
                    }
                }
                document.Save();
                return;
            }

            using ExcelWorkbookWriter writer = ExcelDocument.CreateWriter(format, path, new ExcelWriterOptions { DateSystem = dateSystem });
            switch (route) {
                case "DataTable":
                    writer.WriteTable("Times", CreateEndOfDayTimeTable());
                    break;
                case "DataReader":
                    using (IDataReader reader = CreateEndOfDayTimeTable().CreateDataReader()) {
                        writer.WriteTable("Times", reader);
                    }
                    break;
                default:
                    writer.WriteTable("Times", new[] { CreateEndOfDayTimeRow() });
                    break;
            }
        }

        [Theory]
        [InlineData(1.9999999999988426, false, 1900, 1, 1)]
        [InlineData(45356.99999999999, false, 2024, 3, 5)]
        [InlineData(0.9999999999988426, true, 1904, 1, 1)]
        [InlineData(43894.99999999999, true, 2024, 3, 5)]
        [InlineData(0.9999999999999999, true, 1904, 1, 1)]
        [InlineData(100.99999999999999, true, 1904, 4, 10)]
        public void SerialsJustBeforeMidnightReadOnTheirOwnDay(double serial, bool date1904, int year, int month, int day) {
            ExcelDateSystem dateSystem = date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
            Assert.Equal(new DateTime(year, month, day) + LastMillisecondOfDay, ExcelDateSystemConverter.FromSerial(serial, dateSystem));
        }

        [Fact]
        public void HalfMillisecondSerialRoundsUpInThe1904DateSystem() {
            // 2.5 ms as a double is a little over 2.5 ms; adding 1,462 days first would move it under.
            Assert.Equal(new DateTime(1904, 1, 1, 0, 0, 0, 3), ExcelDateSystemConverter.FromSerial(2.5 / 86_400_000d, ExcelDateSystem.NineteenFour));
        }

        public static IEnumerable<object[]> HalfMillisecondSerials() {
            long[] milliseconds = { 0, 1, 2, 499, 43_199_999, 86_399_998, 86_399_999 };
            foreach (long millisecond in milliseconds) {
                foreach (int day in new[] { 0, 1, 100, 45356 }) yield return new object[] { true, day, millisecond };
                foreach (int day in new[] { 61, 45356 }) yield return new object[] { false, day, millisecond };
            }
        }

        // The time comes from the serial's own fraction, rounded half away from zero, whatever the date system.
        [Theory]
        [MemberData(nameof(HalfMillisecondSerials))]
        public void HalfMillisecondSerialsRoundFromTheirOwnFraction(bool date1904, int day, long millisecond) {
            double serial = day + (millisecond + 0.5d) / 86_400_000d;
            DateTime dayStart = (date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30)).AddDays(day);
            long expected = Math.Min(ExactRoundedMilliseconds(serial - Math.Floor(serial)), 86_399_999L);
            Assert.Equal(dayStart.AddTicks(expected * TimeSpan.TicksPerMillisecond),
                ExcelDateSystemConverter.FromSerial(serial, date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred));
        }

        /// <summary>Rounds fraction x 86,400,000 half up using exact integer arithmetic on the double's bits.</summary>
        private static long ExactRoundedMilliseconds(double fraction) {
            long bits = BitConverter.DoubleToInt64Bits(fraction);
            int exponent = (int)((bits >> 52) & 0x7FF);
            if (exponent == 0) return 0;
            var mantissa = new System.Numerics.BigInteger((bits & 0xF_FFFF_FFFF_FFFFL) | (1L << 52));
            int scale = 1075 - exponent;
            var numerator = mantissa * 86_400_000 * 2 + (System.Numerics.BigInteger.One << scale);
            return (long)(numerator >> (scale + 1));
        }

        public static IEnumerable<object[]> DatesBeforeSerialZero() {
            foreach (bool date1904 in new[] { false, true }) {
                DateTime serialZero = date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 31);
                yield return new object[] { date1904, serialZero.AddTicks(-1) };
                yield return new object[] { date1904, new DateTime(1899, 12, 30).AddTicks(-1) };
                yield return new object[] { date1904, new DateTime(1800, 1, 1, 6, 0, 0).AddTicks(-1) };
            }
        }

        // Dates before serial 0 extend the epoch backwards in both systems; the time floors to the millisecond, so a
        // tick before midnight stays on its own day instead of truncating toward the epoch.
        [Theory]
        [MemberData(nameof(DatesBeforeSerialZero))]
        public void DatesBeforeSerialZeroRoundTripOnTheirOwnDay(bool date1904, DateTime value) {
            ExcelDateSystem dateSystem = date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
            DateTime expected = new DateTime(value.Ticks - value.Ticks % TimeSpan.TicksPerMillisecond);
            double serial = ExcelDateSystemConverter.ToSerial(value, dateSystem);
            Assert.True(serial < 0, $"Serial {serial} should be negative.");
            Assert.Equal(expected, ExcelDateSystemConverter.FromSerial(serial, dateSystem));
            Assert.Equal(expected, ExcelPivotCacheDateCodec.FromSerial(ExcelPivotCacheDateCodec.ToSerial(value, dateSystem), dateSystem));
        }

        // Pivot caches store OLE Automation dates; dates their decoder cannot read are rejected as ToOADate rejects them.
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PivotCacheDatesShareTheOleAutomationRange(bool date1904) {
            ExcelDateSystem dateSystem = date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
            Assert.Throws<OverflowException>(() => ExcelPivotCacheDateCodec.ToSerial(new DateTime(99, 12, 31, 23, 59, 59), dateSystem));
            var first = new DateTime(100, 1, 1, 6, 30, 0);
            Assert.Equal(first, ExcelPivotCacheDateCodec.FromSerial(ExcelPivotCacheDateCodec.ToSerial(first, dateSystem), dateSystem));
        }

        // A negative OLE Automation serial rounds its time from its own fraction, like every other serial: this one is
        // 2.49999... ms after midnight of December 29, 1899, which FromOADate's whole-value rounding made 3 ms.
        [Fact]
        public void NegativeOleAutomationSerialsRoundFromTheirOwnFraction() {
            Assert.Equal(new DateTime(1899, 12, 29, 0, 0, 0, 2),
                ExcelPivotCacheDateCodec.FromSerial(-1.0000000289351851, ExcelDateSystem.NineteenHundred));
        }

        // Displayed fractional seconds and typed reads agree on dates before serial 0 in the 1904 system.
        [Fact]
        public void Negative1904SerialsDisplayTheSameDateTheyReadAs() {
            using var document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "Negative1904Display.xlsx"));
            document.DateSystem = ExcelDateSystem.NineteenFour;
            ExcelSheet sheet = document.AddSheet("Dates");
            sheet.Cell(1, 1).SetValue(-1500.25).SetNumberFormat("yyyy-mm-dd hh:mm:ss.000");
            Assert.Equal(new DateTime(1899, 11, 21, 18, 0, 0), ExcelDateSystemConverter.FromSerial(-1500.25, ExcelDateSystem.NineteenFour));
            Assert.Equal("1899-11-21 18:00:00.000", sheet.Cell(1, 1).GetFormattedText(System.Globalization.CultureInfo.InvariantCulture));
        }

        // Adding the 1904 offset to these serials directly rounds them up to the next whole day.
        [Theory]
        [InlineData(ExcelFileFormat.Xlsx)]
        [InlineData(ExcelFileFormat.Xlsb)]
        [InlineData(ExcelFileFormat.Xls)]
        public void SerialsJustBeforeMidnightKeepTheirDayThroughThe1904Offset(ExcelFileFormat format) {
            string source = Path.Combine(_directoryWithFiles, "EndOfDay1904-source.xlsx");
            string path = Path.Combine(_directoryWithFiles, "EndOfDay1904." + format.ToString().ToLowerInvariant());
            double[] serials = { 0.9999999999999999, 100.99999999999999 };
            DateTime[] expected = { new DateTime(1904, 1, 1) + LastMillisecondOfDay, new DateTime(1904, 4, 10) + LastMillisecondOfDay };
            using (var document = ExcelDocument.Create(source)) {
                document.DateSystem = ExcelDateSystem.NineteenFour;
                ExcelSheet sheet = document.AddSheet("Times");
                sheet.Cell(1, 1).SetValue("When");
                for (int row = 0; row < serials.Length; row++) {
                    sheet.Cell(row + 2, 1).SetValue(serials[row]).SetNumberFormat("yyyy-mm-dd hh:mm:ss");
                }
                document.Save(path, new ExcelSaveOptions { Format = format });
            }

            using (ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path)) {
                foreach (DateTime value in expected) {
                    Assert.True(reader.Read());
                    Assert.Equal(value, reader.GetDateTime(0));
                }
            }

            // Moving the workbook to the 1900 date system keeps each cell on its day.
            using var rebased = new MemoryStream();
            using (var document = ExcelDocument.Load(path)) {
                document.DateSystem = ExcelDateSystem.NineteenHundred;
                document.Save(rebased, new ExcelSaveOptions { Format = ExcelFileFormat.Xlsx });
            }

            rebased.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(rebased, false);
            Cell[] cells = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>()
                .Where(cell => cell.CellReference?.Value is "A2" or "A3").ToArray();
            Assert.Equal(new[] { 1462d, 1562d }, cells.Select(cell => Math.Floor(double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture))));
        }
    }
}
