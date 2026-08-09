using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void Sheets_ExposesAReadOnlyCollectionContract() {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddWorksheet("Data");

        IReadOnlyList<ExcelSheet> sheets = document.Sheets;

        Assert.Single(sheets);
        Assert.Equal("Data", sheets[0].Name);
    }

    [Fact]
    public void RowsAs_UsesThePopulatedRangeWhenNoRangeIsSpecified() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-RowsAs-{Guid.NewGuid():N}.xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "OrderId");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(2, 1, 42);
                sheet.CellValue(2, 2, 165258.24m);
                document.Save();
            }

            using ExcelDocument loaded = ExcelDocument.Load(path);
            TypedSalesRow row = Assert.Single(loaded.GetSheet("Data").RowsAs<TypedSalesRow>());

            Assert.Equal(42, row.OrderId);
            Assert.Equal(165258.24m, row.Amount);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void RangeLessTypedRowsObserveCancellationDuringUsedRangeDiscovery() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "OrderId");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            sheet.RowsAs<TypedSalesRow>(cancellationToken: cancellation.Token).ToArray());
        Assert.Throws<OperationCanceledException>(() =>
            sheet.RowsAs(factory: _ => new PositionalSalesRow(0, 0m), cancellationToken: cancellation.Token).ToArray());

        using var optionsCancellation = new CancellationTokenSource();
        optionsCancellation.Cancel();
        var options = new ExcelReadOptions { CancellationToken = optionsCancellation.Token };
        Assert.Throws<OperationCanceledException>(() =>
            sheet.RowsAs<TypedSalesRow>(options).ToArray());
    }

    [Fact]
    public void RowsAs_ReturnsAnEmptySequenceForAnEmptyWorksheet() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Empty");

        Assert.Empty(sheet.RowsAs<TypedSalesRow>());
    }

    [Fact]
    public void RowsAs_PreservesExcelAliasesAndTypeConverter() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Order Number");
        sheet.CellValue(1, 2, "Amount");
        sheet.CellValue(2, 1, "custom-value");
        sheet.CellValue(2, 2, 12.5m);
        var options = new ExcelReadOptions {
            TypeConverter = static (value, targetType, _) =>
                targetType == typeof(int) && Equals(value, "custom-value")
                    ? (true, 42)
                    : (false, null)
        };

        AliasedSalesRow row = Assert.Single(sheet.RowsAs<AliasedSalesRow>(options));

        Assert.Equal(42, row.OrderId);
        Assert.Equal(12.5m, row.Amount);
    }

#if NET6_0_OR_GREATER
    [Fact]
    public void RowsAs_MapsDateOnlyAndTimeOnlyFromExcelDateCells() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Date");
        sheet.CellValue(1, 2, "Time");
        sheet.CellValue(2, 1, new DateOnly(2026, 8, 6));
        sheet.CellValue(2, 2, new TimeOnly(14, 35, 12));

        DateAndTimeRow row = Assert.Single(sheet.RowsAs<DateAndTimeRow>());

        Assert.Equal(new DateOnly(2026, 8, 6), row.Date);
        Assert.Equal(new TimeOnly(14, 35, 12), row.Time);
    }
#endif

    [Fact]
    public void RowsAs_RedactsTypeConverterFailuresWhenRequested() {
        const string secret = "customer-secret-value";
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "OrderId");
        sheet.CellValue(2, 1, secret);
        var options = new ExcelReadOptions {
            MappingErrorValuePolicy = DataMappingErrorValuePolicy.Redact,
            TypeConverter = (_, _, _) => throw new InvalidOperationException($"failed for {secret}")
        };

        DataMappingException exception = Assert.Throws<DataMappingException>(() =>
            sheet.RowsAs<TypedSalesRow>(options).ToArray());

        Assert.DoesNotContain(secret, exception.ToString(), StringComparison.Ordinal);
        Assert.Contains("Custom converter failed", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowsAs_StrictMappingRejectsUnmappedHeaders() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "OrderId");
        sheet.CellValue(1, 2, "Unexpected");
        sheet.CellValue(2, 1, 42);
        sheet.CellValue(2, 2, "value");

        DataMappingException exception = Assert.Throws<DataMappingException>(() =>
            sheet.RowsAs<TypedSalesRow>(new ExcelReadOptions { StrictTypedMapping = true }).ToArray());

        Assert.Contains("Unexpected", exception.Message);
    }

    [Fact]
    public void RowsAs_ExplicitMapperSupportsUsedAndSpecifiedRanges() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Order Number");
        sheet.CellValue(1, 2, "Amount");
        sheet.CellValue(2, 1, 42);
        sheet.CellValue(2, 2, 12.5m);

        TypedSalesRow usedRangeRow = Assert.Single(sheet.RowsAs<TypedSalesRow>(map => map
            .FromColumn<int>("Order Number", static (row, value) => { row.OrderId = value; return row; })
            .FromColumn<decimal>("Amount", static (row, value) => { row.Amount = value; return row; })));
        TypedSalesRow specifiedRangeRow = Assert.Single(sheet.RowsAs<TypedSalesRow>("A1:B2", map => map
            .FromColumn<int>("Order Number", static (row, value) => { row.OrderId = value; return row; })
            .FromColumn<decimal>("Amount", static (row, value) => { row.Amount = value; return row; })));

        Assert.Equal(42, usedRangeRow.OrderId);
        Assert.Equal(12.5m, usedRangeRow.Amount);
        Assert.Equal(usedRangeRow.OrderId, specifiedRangeRow.OrderId);
        Assert.Equal(usedRangeRow.Amount, specifiedRangeRow.Amount);
    }

    [Fact]
    public void RowsAs_FactorySupportsPositionalRecordsForUsedAndSpecifiedRanges() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Order Number");
        sheet.CellValue(1, 2, "Amount");
        sheet.CellValue(2, 1, 42);
        sheet.CellValue(2, 2, 12.5m);

        PositionalSalesRow usedRangeRow = Assert.Single(sheet.RowsAs(factory: row =>
            new PositionalSalesRow(
                row.GetInt32(row.GetOrdinal("Order Number")),
                row.GetDecimal(row.GetOrdinal("Amount")))));
        PositionalSalesRow specifiedRangeRow = Assert.Single(sheet.RowsAs("A1:B2", factory: row =>
            new PositionalSalesRow(
                row.GetInt32(row.GetOrdinal("Order Number")),
                row.GetDecimal(row.GetOrdinal("Amount")))));

        Assert.Equal(new PositionalSalesRow(42, 12.5m), usedRangeRow);
        Assert.Equal(usedRangeRow, specifiedRangeRow);
    }

    [Fact]
    public void RowsAsParallel_PreservesOrderAcrossAutomaticExplicitAndFactoryMappings() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "OrderId");
        sheet.CellValue(1, 2, "Amount");
        const int rowCount = 1_025;
        for (int index = 0; index < rowCount; index++) {
            sheet.CellValue(index + 2, 1, index);
            sheet.CellValue(index + 2, 2, index + 0.25m);
        }

        var parallel = new ParallelRowMappingOptions {
            MaxDegreeOfParallelism = 4,
            BatchSize = 127
        };
        const string range = "A1:B1026";

        TypedSalesRow[] automatic = sheet.RowsAsParallel<TypedSalesRow>(range, parallel).ToArray();
        TypedSalesRow[] explicitRows = sheet.RowsAsParallel<TypedSalesRow>(range, map => map
            .FromColumn<int>("OrderId", static (row, value) => { row.OrderId = value; return row; })
            .FromColumn<decimal>("Amount", static (row, value) => { row.Amount = value; return row; }), parallel).ToArray();
        PositionalSalesRow[] factoryRows = sheet.RowsAsParallel(range, factory: row =>
            new PositionalSalesRow(row.GetInt32(0), row.GetDecimal(1)), parallelOptions: parallel).ToArray();

        Assert.Equal(Enumerable.Range(0, rowCount), automatic.Select(static row => row.OrderId));
        Assert.Equal(Enumerable.Range(0, rowCount), explicitRows.Select(static row => row.OrderId));
        Assert.Equal(Enumerable.Range(0, rowCount), factoryRows.Select(static row => row.OrderId));
        Assert.Equal(rowCount - 1 + 0.25m, automatic[automatic.Length - 1].Amount);
        Assert.Equal(automatic[automatic.Length - 1].Amount, explicitRows[explicitRows.Length - 1].Amount);
        Assert.Equal(automatic[automatic.Length - 1].Amount, factoryRows[factoryRows.Length - 1].Amount);
    }

    [Fact]
    public void RowsAsParallel_FactoryActuallyRunsConcurrentlyAndPreservesOrder() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "OrderId");
        for (int index = 0; index < 8; index++) {
            sheet.CellValue(index + 2, 1, index);
        }

        using var firstWorkers = new Barrier(2);
        int calls = 0;
        int[] rows = sheet.RowsAsParallel(
            "A1:A9",
            factory: record => {
                if (Interlocked.Increment(ref calls) <= 2) {
                    Assert.True(firstWorkers.SignalAndWait(TimeSpan.FromSeconds(10)));
                }
                return record.GetInt32(0);
            },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 1
            }).ToArray();

        Assert.Equal(Enumerable.Range(0, 8), rows);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsm")]
    [InlineData(".xltx")]
    [InlineData(".xltm")]
    [InlineData(".xlam")]
    [InlineData(".xlsb")]
    [InlineData(".xls")]
    public void OpenDataReader_RowsAsParallelPreservesOrderAcrossSupportedPathFormats(string extension) {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO-ParallelFormat-{Guid.NewGuid():N}{extension}");
        const int rowCount = 257;
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "OrderId");
                sheet.CellValue(1, 2, "Amount");
                for (int index = 0; index < rowCount; index++) {
                    sheet.CellValue(index + 2, 1, index);
                    sheet.CellValue(index + 2, 2, index + 0.25m);
                }
                document.Save();
            }

            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
                path,
                new ExcelReadOptions {
                    SheetName = "Data",
                    InferSchema = true
                });
            TypedSalesRow[] rows = reader.RowsAsParallel<TypedSalesRow>(
                new ParallelRowMappingOptions {
                    MaxDegreeOfParallelism = 4,
                    BatchSize = 31
                }).ToArray();

            Assert.Equal(Enumerable.Range(0, rowCount), rows.Select(static row => row.OrderId));
            Assert.Equal(rowCount - 1 + 0.25m, rows[rowCount - 1].Amount);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void RowsAsParallel_ObservesCancellationBeforeReading() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "OrderId");
        sheet.CellValue(2, 1, 42);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => sheet.RowsAsParallel<TypedSalesRow>(
            new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2 },
            cancellationToken: cancellation.Token).ToArray());
    }

    [Fact]
    public void Reader_ForcedParallelFastPath_PreservesTypedMappingContractsAboveItsActivationBoundary() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-ParallelFastContracts-{Guid.NewGuid():N}.xlsx");
        const int dataRows = 4_097;
        DateTime date = new(2026, 8, 8, 12, 30, 0, DateTimeKind.Unspecified);
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Order Number");
                sheet.CellValue(1, 2, " Amount ");
                sheet.CellValue(1, 3, "When");
                sheet.CellValue(1, 4, "Unexpected");
                for (int index = 0; index < dataRows; index++) {
                    int row = index + 2;
                    sheet.CellValue(row, 1, index + 1);
                    if (index != 0) sheet.CellValue(row, 2, index + 0.25m);
                    sheet.CellValue(row, 3, date.AddDays(index));
                    sheet.CellValue(row, 4, index == 0 ? "customer-secret-value" : index);
                }
                document.Save();
            }

            var options = new ExcelReadOptions {
                NormalizeHeaders = true,
                TreatDatesUsingNumberFormat = true
            };
            options.Execution.MaxDegreeOfParallelism = 4;
            using (ExcelDocumentReader reader = ExcelDocumentReader.Open(path, options)) {
                LargeParallelMappedRow[] rows = reader.GetSheet("Data")
                    .ReadObjects<LargeParallelMappedRow>(
                        $"A1:D{dataRows + 1}",
                        ExcelExecutionMode.Parallel)
                    .ToArray();

                Assert.Equal(dataRows, rows.Length);
                Assert.Equal(1, rows[0].OrderId);
                Assert.Equal(0m, rows[0].Amount);
                Assert.Equal(date.ToOADate(), rows[0].When, precision: 8);
                Assert.Equal(dataRows, rows[dataRows - 1].OrderId);
                Assert.Equal(dataRows - 1 + 0.25m, rows[dataRows - 1].Amount);
            }

            using (ExcelDocumentReader strictReader = ExcelDocumentReader.Open(
                       path,
                       new ExcelReadOptions { StrictTypedMapping = true })) {
                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => strictReader
                    .GetSheet("Data")
                    .ReadObjects<LargeParallelMappedRow>(
                        $"A1:D{dataRows + 1}",
                        ExcelExecutionMode.Parallel)
                    .ToArray());
                Assert.Contains("Unexpected", exception.Message, StringComparison.Ordinal);
            }

        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private sealed class TypedSalesRow {
        public int OrderId { get; set; }
        public decimal Amount { get; set; }
    }

    private sealed class AliasedSalesRow {
        [ExcelColumn("Order Number")]
        public int OrderId { get; set; }
        public decimal Amount { get; set; }
    }

    private sealed class LargeParallelMappedRow {
        [ExcelColumn("Order Number")]
        public int OrderId { get; set; }

        public decimal Amount { get; set; }

        public double When { get; set; }
    }

#if NET6_0_OR_GREATER
    private sealed class DateAndTimeRow {
        public DateOnly Date { get; set; }
        public TimeOnly Time { get; set; }
    }
#endif

    private sealed record PositionalSalesRow(int OrderId, decimal Amount);
}
