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
            // A calendar date read as a duration is its serial's day count, not the distance from December 30, 1899.
            AssertTemporalMapping(format, dateSystem, when, when - (date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30)));

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
            foreach (TemporalMappingRoute<TTarget> route in TemporalMappingRoutes<TTarget>(path, format)) {
                TTarget actual = route.Read();
                Assert.True(Equals(expected, actual), $"{route.Name} on {format} {dateSystem}: expected {expected}, got {actual}.");
            }
        }

        private void AssertTemporalMappingFails<TTarget>(ExcelFileFormat format, ExcelDateSystem dateSystem) {
            string path = WriteTemporalMappingWorkbook(format, dateSystem, "not a time");
            foreach (TemporalMappingRoute<TTarget> route in TemporalMappingRoutes<TTarget>(path, format)) {
                Exception? error = Record.Exception(() => route.Read());
                Assert.True(error?.GetType() == route.Failure,
                    $"{route.Name} on {format} {dateSystem}: expected {route.Failure.Name}, got {(error == null ? "a value" : error.GetType().Name)}.");
            }
        }

        private sealed record TemporalMappingRoute<TTarget>(string Name, Func<TTarget> Read, Type Failure);

        private static IEnumerable<TemporalMappingRoute<TTarget>> TemporalMappingRoutes<TTarget>(string path, ExcelFileFormat format) {
            // Object mapping reports a value that is not a date, time or duration as a mapping failure. GetFieldValue
            // for DateTime is the data reader's own GetDateTime, which fails for text as data readers do.
            Type mapping = typeof(DataMappingException);
            Type getter = (Nullable.GetUnderlyingType(typeof(TTarget)) ?? typeof(TTarget)) == typeof(DateTime) ? typeof(FormatException) : mapping;
            foreach (bool numericAsDecimal in new[] { false, true }) {
                var options = new ExcelReaderOptions { NumericAsDecimal = numericAsDecimal };
                string suffix = numericAsDecimal ? " with decimal numbers" : string.Empty;
                if (format == ExcelFileFormat.Xlsx) {
                    yield return new("ReadObjects" + suffix, () => {
                        using var reader = ExcelDocumentReader.Open(path, options);
                        return reader.GetSheet("Data").ReadObjects<TemporalMappingTarget<TTarget>>("A1:A2").Single().Value;
                    }, mapping);
                    yield return new("ReadObjectsStream" + suffix, () => {
                        using var reader = ExcelDocumentReader.Open(path, options);
                        return reader.GetSheet("Data").ReadObjectsStream<TemporalMappingTarget<TTarget>>("A1:A2").Single().Value;
                    }, mapping);
                    yield return new("ReadColumnAs" + suffix, () => {
                        using var reader = ExcelDocumentReader.Open(path, options);
                        return reader.GetSheet("Data").ReadColumnAs<TTarget>("A2:A2").Single();
                    }, mapping);
                    yield return new("ReadRangeAs" + suffix, () => {
                        using var reader = ExcelDocumentReader.Open(path, options);
                        return reader.GetSheet("Data").ReadRangeAs<TTarget>("A2:A2")[0, 0];
                    }, mapping);
                }

                yield return new("RowsAs" + suffix, () => {
                    using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path, options);
                    return reader.RowsAs<TemporalMappingTarget<TTarget>>().Single().Value;
                }, mapping);
                yield return new("GetFieldValue" + suffix, () => {
                    using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path, options);
                    Assert.True(reader.Read());
                    return reader.GetFieldValue<TTarget>(0);
                }, getter);
            }

            yield return new("RowsAs with an explicit mapper", () => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path);
                return reader.RowsAs<TemporalMappingTarget<TTarget>>(mapper =>
                    mapper.FromColumn<TTarget>("Value", (target, value) => { target.Value = value; return target; })).Single().Value;
            }, mapping);
            yield return new("RowsAsParallel", () => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path);
                return reader.RowsAsParallel<TemporalMappingTarget<TTarget>>(
                    new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 }).Single().Value;
            }, mapping);
            yield return new("RowsAs with schema inference", () => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { InferSchema = true, SchemaSampleRows = 8 });
                return reader.RowsAs<TemporalMappingTarget<TTarget>>().Single().Value;
            }, mapping);
            yield return new("GetFieldValue with schema inference", () => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path, new ExcelReaderOptions { InferSchema = true, SchemaSampleRows = 8 });
                Assert.True(reader.Read());
                return reader.GetFieldValue<TTarget>(0);
            }, getter);
            yield return new("ExcelSheet.RowsAs", () => {
                using var document = ExcelDocument.Load(path);
                return document.Sheets[0].RowsAs<TemporalMappingTarget<TTarget>>().Single().Value;
            }, mapping);
        }
    }
}
