using System.Data;
using OfficeIMO.Data;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public sealed class TemporalMappingSource<T> {
            public T Value { get; set; } = default!;
        }

        public sealed class TemporalMappingTarget<T> {
            public T Value { get; set; } = default!;
        }

        public static IEnumerable<object[]> TemporalMappingWorkbooks() {
            foreach (bool date1904 in new[] { false, true }) {
                foreach (ExcelFileFormat format in new[] { ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb, ExcelFileFormat.Xls }) {
                    yield return new object[] { format, date1904 };
                }
            }
        }

        [Theory]
        [MemberData(nameof(TemporalMappingWorkbooks))]
        public void TemporalCellsMapToEveryDateAndTimeDestination(ExcelFileFormat format, bool date1904) {
            ExcelDateSystem dateSystem = date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
            var elapsed = new TimeSpan(1, 2, 30, 0);
            AssertTemporalMapping(format, dateSystem, elapsed, elapsed);
            AssertTemporalMapping<TimeSpan, TimeSpan?>(format, dateSystem, elapsed, elapsed);
            AssertTemporalMapping(format, dateSystem, TimeSpan.FromHours(-6), TimeSpan.FromHours(-6));

            var when = new DateTime(2024, 3, 5, 10, 30, 0);
            AssertTemporalMapping(format, dateSystem, when, new DateTimeOffset(when));
            AssertTemporalMapping<DateTime, DateTimeOffset?>(format, dateSystem, when, new DateTimeOffset(when));

            // A plain number read as a date is a serial in the workbook's date system.
            DateTime epoch = date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
            AssertTemporalMapping(format, dateSystem, 45356.5, epoch.AddDays(45356.5));
#if NET6_0_OR_GREATER
            AssertTemporalMapping(format, dateSystem, new TimeOnly(10, 30), new TimeOnly(10, 30));
            AssertTemporalMapping(format, dateSystem, new TimeOnly(10, 30), new TimeSpan(10, 30, 0));
            AssertTemporalMapping<DateOnly, DateOnly?>(format, dateSystem, new DateOnly(2024, 3, 5), new DateOnly(2024, 3, 5));
            AssertTemporalMapping(format, dateSystem, 45356.5, DateOnly.FromDateTime(epoch.AddDays(45356.5)));
#endif
        }

        [Theory]
        [MemberData(nameof(TemporalMappingWorkbooks))]
        public void TextThatIsNotADateOrTimeFailsInsteadOfMappingADefault(ExcelFileFormat format, bool date1904) {
            ExcelDateSystem dateSystem = date1904 ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred;
            AssertTemporalMappingFails<TimeSpan>(format, dateSystem);
            AssertTemporalMappingFails<TimeSpan?>(format, dateSystem);
            AssertTemporalMappingFails<DateTimeOffset>(format, dateSystem);
            AssertTemporalMappingFails<DateTime>(format, dateSystem);
#if NET6_0_OR_GREATER
            AssertTemporalMappingFails<DateOnly>(format, dateSystem);
            AssertTemporalMappingFails<TimeOnly?>(format, dateSystem);
#endif
        }

        private string WriteTemporalMappingWorkbook<TSource>(ExcelFileFormat format, ExcelDateSystem dateSystem, TSource value) {
            string path = Path.Combine(_directoryWithFiles, $"TemporalMapping-{Guid.NewGuid():N}.{format.ToString().ToLowerInvariant()}");
            using var writer = ExcelDocument.CreateWriter(format, path, new ExcelWriterOptions { DateSystem = dateSystem });
            writer.WriteTable("Data", new[] { new TemporalMappingSource<TSource> { Value = value } });
            return path;
        }

        private void AssertTemporalMapping<TSource, TTarget>(ExcelFileFormat format, ExcelDateSystem dateSystem, TSource value, TTarget expected) {
            string path = WriteTemporalMappingWorkbook(format, dateSystem, value);
            foreach ((string route, Func<TTarget> read) in TemporalMappingRoutes<TTarget>(path, format)) {
                Assert.True(Equals(expected, read()), $"{route} on {format} {dateSystem}: expected {expected}, got {read()}.");
            }
        }

        private void AssertTemporalMappingFails<TTarget>(ExcelFileFormat format, ExcelDateSystem dateSystem) {
            string path = WriteTemporalMappingWorkbook(format, dateSystem, "not a time");
            foreach ((string route, Func<TTarget> read) in TemporalMappingRoutes<TTarget>(path, format)) {
                Exception? error = Record.Exception(() => read());
                Assert.True(error is DataMappingException or InvalidCastException or FormatException, $"{route} on {format} {dateSystem} returned {(error == null ? "a value" : error.GetType().Name)}.");
            }
        }

        private static IEnumerable<(string Route, Func<TTarget> Read)> TemporalMappingRoutes<TTarget>(string path, ExcelFileFormat format) {
            if (format == ExcelFileFormat.Xlsx) {
                yield return ("ReadObjects", () => {
                    using var reader = ExcelDocumentReader.Open(path);
                    return reader.GetSheet("Data").ReadObjects<TemporalMappingTarget<TTarget>>("A1:A2").Single().Value;
                });
                yield return ("ReadObjectsStream", () => {
                    using var reader = ExcelDocumentReader.Open(path);
                    return reader.GetSheet("Data").ReadObjectsStream<TemporalMappingTarget<TTarget>>("A1:A2").Single().Value;
                });
                yield return ("ReadColumnAs", () => {
                    using var reader = ExcelDocumentReader.Open(path);
                    return reader.GetSheet("Data").ReadColumnAs<TTarget>("A2:A2").Single();
                });
            }

            yield return ("RowsAs", () => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path);
                return reader.RowsAs<TemporalMappingTarget<TTarget>>().Single().Value;
            });
            yield return ("GetFieldValue", () => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path);
                Assert.True(reader.Read());
                return reader.GetFieldValue<TTarget>(0);
            });
            yield return ("RowsAs with an explicit mapper", () => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path);
                return reader.RowsAs<TemporalMappingTarget<TTarget>>(mapper =>
                    mapper.FromColumn<TTarget>("Value", (target, value) => { target.Value = value; return target; })).Single().Value;
            });
            yield return ("RowsAsParallel", () => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path);
                return reader.RowsAsParallel<TemporalMappingTarget<TTarget>>(
                    new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 }).Single().Value;
            });
            yield return ("RowsAs with schema inference", () => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { InferSchema = true, SchemaSampleRows = 8 });
                return reader.RowsAs<TemporalMappingTarget<TTarget>>().Single().Value;
            });
            yield return ("GetFieldValue with schema inference", () => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { InferSchema = true, SchemaSampleRows = 8 });
                Assert.True(reader.Read());
                return reader.GetFieldValue<TTarget>(0);
            });
            yield return ("ExcelSheet.RowsAs", () => {
                using var document = ExcelDocument.Load(path);
                return document.Sheets[0].RowsAs<TemporalMappingTarget<TTarget>>().Single().Value;
            });
        }
    }
}
