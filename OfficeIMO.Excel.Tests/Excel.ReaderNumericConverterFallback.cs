using OfficeIMO.Data;
using OfficeIMO.Excel;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    private sealed class NumericFallbackRow {
        public double? Value { get; set; }
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred, "[h]:mm")]
    [InlineData(ExcelDateSystem.NineteenFour, "[h]:mm")]
    [InlineData(ExcelDateSystem.NineteenHundred, "yyyy-mm-dd")]
    [InlineData(ExcelDateSystem.NineteenFour, "yyyy-mm-dd")]
    public void Reader_NumericConverterFallback_PreservesSerialAndConverterPrecedence(ExcelDateSystem system, string format) {
        string path = Path.Combine(_directoryWithFiles, $"NumericFallback{system}{format.Replace(':', '-').Replace('/', '-')}.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            document.DateSystem = system;
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            sheet.CellValue(2, 1, 1.5d);
            sheet.CellFormula(3, 1, "1+0.5");
            sheet.ColumnStyleByHeader("Value").NumberFormat(format);
            Assert.Equal(1, document.RecalculateSupportedFormulas());
            document.Save();
        }
        // Constant and cached formula values must follow the same contract.
        for (int mode = 0; mode < 8; mode++) {
            for (int api = 0; api < 7; api++) {
                int cellCalls = 0, typeCalls = 0;
                var options = new ExcelReadOptions { SheetName = "Data" };
                if (mode != 1) {
                    options.CellValueConverter = context => {
                        if (context.RawText != "1.5") return ExcelCellValue.NotHandled;
                        Interlocked.Increment(ref cellCalls);
                        return mode is 3 or 7 ? new ExcelCellValue(17d)
                            : mode == 4 ? new ExcelCellValue(null)
                            : ExcelCellValue.NotHandled;
                    };
                }
                if (mode != 0) {
                    options.TypeConverter = (value, type, _) => {
                        Assert.Equal(typeof(double), type);
                        Interlocked.Increment(ref typeCalls);
                        if (mode is 5 or 7) return (true, 29d);
                        if (mode == 6) return (true, null);
                        Assert.True(value is DateTime || Equals(value, 17d) || (api >= 4 && Equals(value, 1.5d)),
                            $"Unexpected converter input {value} for mode {mode}, API {api}.");
                        return (false, null);
                    };
                }
                NumericFallbackRow[] rows;
                if (api < 4) {
                    using var reader = ExcelDocument.OpenDataReader(path, options);
                    Action<RowMapper<NumericFallbackRow>> configure = mapper => mapper.FromColumn<double?>("Value", (row, value) => { row.Value = value; return row; });
                    var parallel = new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 };
                    rows = api switch {
                        0 => reader.RowsAs<NumericFallbackRow>().ToArray(),
                        1 => reader.RowsAsParallel<NumericFallbackRow>(parallel).ToArray(),
                        2 => reader.RowsAs(configure).ToArray(),
                        _ => reader.RowsAsParallel(configure, parallel).ToArray()
                    };
                } else {
                    using var reader = ExcelDocumentReader.Open(path, options);
                    var sheet = reader.GetSheet("Data");
                    rows = api == 4 ? sheet.ReadObjectsStream<NumericFallbackRow>("A1:A3").ToArray()
                        : sheet.ReadObjects<NumericFallbackRow>("A1:A3", api == 5 ? ExcelExecutionMode.Sequential : ExcelExecutionMode.Parallel).ToArray();
                }
                Assert.Equal(2, rows.Length);
                double? expected = mode == 3 ? 17d : mode is 4 or 6 ? null : mode is 5 or 7 ? 29d : 1.5d;
                Assert.All(rows, row => Assert.Equal(expected, row.Value));
                Assert.Equal(options.CellValueConverter == null ? 0 : 2, cellCalls);
                Assert.Equal(options.TypeConverter == null || mode == 4 ? 0 : 2, typeCalls);
            }
        }
    }

    [Fact]
    public void Reader_NumericConverterFallback_DoesNotReplaceHandledDateOrSuppressConverterFailure() {
        string path = Path.Combine(_directoryWithFiles, "NumericConverterFailure.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            sheet.CellValue(2, 1, 1.5d);
            sheet.ColumnStyleByHeader("Value").NumberFormat("[h]:mm");
            document.Save();
        }
        foreach (bool parallel in new[] { false, true }) {
            foreach (bool handledCell in new[] { false, true }) {
                int converterCalls = 0;
                var options = new ExcelReadOptions {
                    SheetName = "Data",
                    MappingErrorValuePolicy = DataMappingErrorValuePolicy.Redact,
                    CellValueConverter = context => handledCell && context.RawText == "1.5"
                        ? new ExcelCellValue(new DateTime(2024, 1, 1)) : ExcelCellValue.NotHandled,
                    TypeConverter = (_, _, _) => {
                        Interlocked.Increment(ref converterCalls);
                        if (!handledCell) throw new InvalidOperationException("private-converter-detail");
                        return (false, null);
                    }
                };
                using var reader = ExcelDocument.OpenDataReader(path, options);
                var error = Assert.Throws<DataMappingException>(() => (parallel
                    ? reader.RowsAsParallel<NumericFallbackRow>(new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 })
                    : reader.RowsAs<NumericFallbackRow>()).ToArray());
                Assert.Equal(1, converterCalls);
                Assert.DoesNotContain("private-converter-detail", error.Message);
            }
        }
    }

    [Theory]
    [InlineData(ExcelDateSystem.NineteenHundred, "[h]:mm")]
    [InlineData(ExcelDateSystem.NineteenFour, "[h]:mm")]
    [InlineData(ExcelDateSystem.NineteenHundred, "yyyy-mm-dd")]
    [InlineData(ExcelDateSystem.NineteenFour, "yyyy-mm-dd")]
    public void Reader_LargeParallelMapping_PreservesDurationAndCalendarSerial60(ExcelDateSystem system, string format) {
        string path = Path.Combine(_directoryWithFiles, $"LargeNumericSerial{system}{format.Replace(':', '-')}.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            document.DateSystem = system;
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Value");
            for (int row = 2; row <= 4100; row++) sheet.CellValue(row, 1, row % 2 == 0 ? 1.5d : 60.5d);
            sheet.ColumnStyleByHeader("Value").NumberFormat(format);
            document.Save();
        }
        using var reader = ExcelDocumentReader.Open(path);
        var rows = reader.GetSheet("Data").ReadObjects<NumericFallbackRow>("A1:A4100", ExcelExecutionMode.Parallel).ToArray();
        Assert.Equal(4099, rows.Length);
        for (int i = 0; i < rows.Length; i++) Assert.Equal(i % 2 == 0 ? 1.5d : 60.5d, rows[i].Value);
    }
}
