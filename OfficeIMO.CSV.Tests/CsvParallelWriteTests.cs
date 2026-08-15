#nullable enable

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvParallelWriteTests
{
    [Fact]
    public void WriteDataReaderParallel_MatchesSequentialOutputAcrossOrderedBatches()
    {
        using DataTable table = CreateMixedTable(rowCount: 19);
        var options = new CsvSaveOptions { NewLine = "\n" };

        string sequential = WriteSequential(table, options);
        string parallel = WriteParallel(
            table,
            options,
            new CsvWriteParallelOptions { MaxDegreeOfParallelism = 3, BatchSize = 4 });

        Assert.Equal(sequential, parallel);
        CsvDocument parsed = CsvDocument.Parse(parallel);
        CsvRow[] rows = parsed.AsEnumerable().ToArray();
        Assert.Equal(19, rows.Length);
        Assert.Equal("Name 18", rows[18].AsString("Name"));
    }

    [Fact]
    public void WriteDataReaderParallel_MatchesSequentialFormattingOptions()
    {
        using var table = new DataTable("Formatting");
        table.Columns.Add("Name", typeof(object));
        table.Columns.Add("When", typeof(object));
        table.Columns.Add("Value", typeof(object));
        table.Rows.Add("=SUM(A1:A2)", new DateTimeOffset(2026, 8, 10, 12, 30, 0, TimeSpan.FromHours(2)), 12.5m);
        table.Rows.Add("A||B\nC", DBNull.Value, true);
        table.Rows.Add(" spaced ", new DateTime(2026, 8, 10, 10, 45, 0, DateTimeKind.Utc), -7);
        var options = new CsvSaveOptions
        {
            DelimiterText = "||",
            NewLine = "\r\n",
            NullValue = "<null>",
            DateTimeFormat = "O",
            UseUtc = true,
            FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape,
            QuoteMode = CsvQuoteMode.Always
        };

        Assert.Equal(
            WriteSequential(table, options),
            WriteParallel(
                table,
                options,
                new CsvWriteParallelOptions { MaxDegreeOfParallelism = 2, BatchSize = 2 }));
    }

    [Fact]
    public void WriteDataReaderParallel_ConsumesReaderOnOneThread()
    {
        using DataTable table = CreateMixedTable(rowCount: 23);
        using var reader = new ThreadTrackingDataReader(table.CreateDataReader());
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        CsvDocument.WriteDataReaderParallel(
            writer,
            reader,
            new CsvSaveOptions { NewLine = "\n" },
            new CsvWriteParallelOptions { MaxDegreeOfParallelism = 4, BatchSize = 3 });

        Assert.Single(reader.AccessThreadIds);
        Assert.Equal(23, CsvDocument.Parse(writer.ToString()).AsEnumerable().Count());
    }

    [Fact]
    public void WriteDataReaderParallel_DegreeOneUsesSequentialContract()
    {
        using DataTable table = CreateMixedTable(rowCount: 7);
        var options = new CsvSaveOptions { NewLine = "\n", QuoteMode = CsvQuoteMode.Always };

        Assert.Equal(
            WriteSequential(table, options),
            WriteParallel(
                table,
                options,
                new CsvWriteParallelOptions { MaxDegreeOfParallelism = 1, BatchSize = 1 }));
    }

    [Fact]
    public void WriteDataReaderParallel_FallsBackWhenGetValuesIsUnsupported()
    {
        object?[][] rows =
        {
            new object?[] { 1, "Alpha", null },
            new object?[] { 2, "Beta,quoted", true }
        };
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Id", "Name", "Enabled" },
            rows);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        CsvDocument.WriteDataReaderParallel(
            writer,
            reader,
            new CsvSaveOptions { NewLine = "\n" },
            new CsvWriteParallelOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 });

        Assert.Equal("Id,Name,Enabled\n1,Alpha,\n2,\"Beta,quoted\",True\n", writer.ToString());
        Assert.Equal(6, reader.GetValueCallCount);
    }

    [Fact]
    public void WriteDataReaderParallel_SnapshotsUnsafeProviderValuesSequentially()
    {
        var shared = new MutableFormattable();
        using var reader = new ThrowingGetValuesDataReader(
            ["Value"],
            [[shared], [shared], [shared]],
            rowIndex => shared.Value = ((char)('A' + rowIndex)).ToString(CultureInfo.InvariantCulture));
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        CsvDocument.WriteDataReaderParallel(
            writer,
            reader,
            new CsvSaveOptions { NewLine = "\n" },
            new CsvWriteParallelOptions { MaxDegreeOfParallelism = 2, BatchSize = 3 });

        Assert.Equal("Value\nA\nB\nC\n", writer.ToString());
    }

    [Fact]
    public void WriteDataReaderParallel_RejectsWideUnsafeProviderBeforeSnapshotPlanning()
    {
        var shared = new MutableFormattable { Value = "safe" };
        using var reader = new ThrowingGetValuesDataReader(
            ["First", "Second", "Third"],
            [[1, 2, shared]]);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            CsvDocument.WriteDataReaderParallel(
                writer,
                reader,
                new CsvSaveOptions { NewLine = "\n" },
                new CsvWriteParallelOptions
                {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 4096,
                    MaximumBufferedCellsPerBatch = 2
                }));

        Assert.Contains("per-batch cell budget", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, reader.GetFieldTypeCallCount);
        Assert.Equal(0, reader.GetValueCallCount);
        Assert.Equal(string.Empty, writer.ToString());
    }

    [Fact]
    public void WriteDataReaderParallel_EmptyReaderMatchesSequentialHeader()
    {
        using var table = new DataTable("Empty");
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        var options = new CsvSaveOptions { NewLine = "\n" };

        Assert.Equal(
            WriteSequential(table, options),
            WriteParallel(
                table,
                options,
                new CsvWriteParallelOptions { MaxDegreeOfParallelism = 3, BatchSize = 2 }));
    }

    [Fact]
    public void WriteDataReaderParallel_BoundsBatchRowsByReaderWidth()
    {
        var options = new CsvWriteParallelOptions
        {
            BatchSize = 4096,
            MaximumBufferedCellsPerBatch = 200
        };

        Assert.Equal(2, options.GetBatchSize(fieldCount: 100));
    }

    [Fact]
    public void WriteDataReaderParallel_RejectsSchemasWiderThanTheCellBudgetBeforeReading()
    {
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "A", "B", "C" },
            [new object?[] { 1, 2, 3 }]);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            CsvDocument.WriteDataReaderParallel(
                writer,
                reader,
                new CsvSaveOptions { NewLine = "\n" },
                new CsvWriteParallelOptions
                {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 4096,
                    MaximumBufferedCellsPerBatch = 2
                }));

        Assert.Contains("per-batch cell budget", exception.Message, StringComparison.Ordinal);
        Assert.Equal(0, reader.GetValueCallCount);
        Assert.Equal(0, reader.GetFieldTypeCallCount);
        Assert.Equal(string.Empty, writer.ToString());
    }

    [Theory]
    [InlineData(0, 4)]
    [InlineData(-1, 4)]
    [InlineData(2, 0)]
    [InlineData(2, -1)]
    public void WriteDataReaderParallel_InvalidLimitsDoNotMutateDestination(int degree, int batchSize)
    {
        using DataTable table = CreateMixedTable(rowCount: 1);
        using DataTableReader reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        Assert.Throws<ArgumentOutOfRangeException>(() => CsvDocument.WriteDataReaderParallel(
            writer,
            reader,
            new CsvSaveOptions { NewLine = "\n" },
            new CsvWriteParallelOptions
            {
                MaxDegreeOfParallelism = degree,
                BatchSize = batchSize
            }));
        Assert.Equal(string.Empty, writer.ToString());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void WriteDataReaderParallel_InvalidCellBudgetIsRejectedOnSequentialPath(int maximumBufferedCells)
    {
        using DataTable table = CreateMixedTable(rowCount: 1);
        using DataTableReader reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            CsvDocument.WriteDataReaderParallel(
                writer,
                reader,
                new CsvSaveOptions { NewLine = "\n" },
                new CsvWriteParallelOptions
                {
                    MaxDegreeOfParallelism = 1,
                    MaximumBufferedCellsPerBatch = maximumBufferedCells
                }));

        Assert.Equal(nameof(CsvWriteParallelOptions.MaximumBufferedCellsPerBatch), exception.ParamName);
        Assert.Equal(string.Empty, writer.ToString());
    }

    [Fact]
    public void WriteDataReaderParallel_UnwrapsSingleFormattingFailure()
    {
        using var table = new DataTable("Failure");
        table.Columns.Add("Value", typeof(object));
        table.Rows.Add(new ThrowingFormattable());
        using DataTableReader reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            CsvDocument.WriteDataReaderParallel(
                writer,
                reader,
                new CsvSaveOptions { IncludeHeader = false, NewLine = "\n" },
                new CsvWriteParallelOptions { MaxDegreeOfParallelism = 4, BatchSize = 1 }));

        Assert.Equal("format failed", exception.Message);
    }

    [Fact]
    public void WriteDataReaderParallel_FormattingFailurePrecedesConcurrentReadFailure()
    {
        using var formattingFailed = new ManualResetEventSlim();
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Value" },
            new object?[][]
            {
                new object?[] { new SignalingThrowingFormattable(formattingFailed) },
                new object?[] { "same batch" },
                new object?[] { "read ahead" }
            },
            rowIndex =>
            {
                if (rowIndex != 2)
                {
                    return;
                }

                if (!formattingFailed.Wait(TimeSpan.FromSeconds(5)))
                {
                    throw new TimeoutException("Formatter did not run before the read-ahead failure.");
                }

                throw new InvalidDataException("read failed later");
            });
        using var writer = new StringWriter(CultureInfo.InvariantCulture);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            CsvDocument.WriteDataReaderParallel(
                writer,
                reader,
                new CsvSaveOptions { IncludeHeader = false, NewLine = "\n" },
                new CsvWriteParallelOptions { MaxDegreeOfParallelism = 2, BatchSize = 2 }));

        Assert.Equal("format failed first", exception.Message);
    }

    [Fact]
    public void WriteDataReaderParallel_CapsWorkersToUsefulBatchConcurrency()
    {
        using DataTable table = CreateMixedTable(rowCount: 1);
        string sequential = WriteSequential(table, new CsvSaveOptions { NewLine = "\n" });

        string parallel = WriteParallel(
            table,
            new CsvSaveOptions { NewLine = "\n" },
            new CsvWriteParallelOptions
            {
                MaxDegreeOfParallelism = int.MaxValue,
                BatchSize = 1
            });

        Assert.Equal(sequential, parallel);
    }

    [Fact]
    public void WriteDataReaderParallel_CancellationDoesNotReplacePathDestination()
    {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.ParallelCancel.{Guid.NewGuid():N}.csv");
        const string original = "Name\nOriginal\n";
        File.WriteAllText(path, original);
        using var cancellation = new CancellationTokenSource();
        using var table = new DataTable("Cancellation");
        table.Columns.Add("Value", typeof(object));
        table.Rows.Add(new CancelingFormattable(cancellation));
        table.Rows.Add("still staged");
        table.Rows.Add("read ahead");
        using DataTableReader reader = table.CreateDataReader();

        try
        {
            Assert.Throws<OperationCanceledException>(() => CsvDocument.WriteDataReaderParallel(
                path,
                reader,
                new CsvSaveOptions { NewLine = "\n" },
                new CsvWriteParallelOptions { MaxDegreeOfParallelism = 2, BatchSize = 2 },
                cancellation.Token));

            Assert.Equal(original, File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void WriteDataReaderParallel_ReadFailureDoesNotReplacePathDestination()
    {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.ParallelReadFailure.{Guid.NewGuid():N}.csv");
        const string original = "Name\nOriginal\n";
        File.WriteAllText(path, original);
        using var reader = new ThrowingGetValuesDataReader(
            new[] { "Id", "Name" },
            new object?[][]
            {
                new object?[] { 1, "Alpha" },
                new object?[] { 2, "Beta" },
                new object?[] { 3, "Gamma" }
            },
            rowIndex =>
            {
                if (rowIndex == 2) throw new InvalidOperationException("read failed");
            });

        try
        {
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                CsvDocument.WriteDataReaderParallel(
                    path,
                    reader,
                    new CsvSaveOptions { NewLine = "\n" },
                    new CsvWriteParallelOptions { MaxDegreeOfParallelism = 2, BatchSize = 2 }));

            Assert.Equal("read failed", exception.Message);
            Assert.Equal(original, File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void WriteDataReaderParallel_StreamOverloadLeavesDestinationOpen()
    {
        using DataTable table = CreateMixedTable(rowCount: 5);
        using DataTableReader reader = table.CreateDataReader();
        using var destination = new MemoryStream();

        CsvDocument.WriteDataReaderParallel(
            destination,
            reader,
            new CsvSaveOptions { NewLine = "\n" },
            new CsvWriteParallelOptions { MaxDegreeOfParallelism = 2, BatchSize = 2 });

        Assert.True(destination.CanWrite);
        Assert.Equal(5, CsvDocument.Parse(System.Text.Encoding.UTF8.GetString(destination.ToArray())).AsEnumerable().Count());
    }

    private static DataTable CreateMixedTable(int rowCount)
    {
        var table = new DataTable("Mixed") { Locale = CultureInfo.InvariantCulture };
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("When", typeof(DateTime));
        table.Columns.Add("Amount", typeof(decimal));
        table.Columns.Add("Enabled", typeof(bool));
        table.Columns.Add("Notes", typeof(object));
        for (int index = 0; index < rowCount; index++)
        {
            table.Rows.Add(
                index,
                "Name " + index,
                new DateTime(2026, 8, 10, 8, 0, 0, DateTimeKind.Utc).AddMinutes(index),
                index + 0.25m,
                (index & 1) == 0,
                index % 3 == 0 ? DBNull.Value : $"row {index}, quoted\ntext");
        }

        return table;
    }

    private static string WriteSequential(DataTable table, CsvSaveOptions options)
    {
        using DataTableReader reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteDataReader(writer, reader, options);
        return writer.ToString();
    }

    private static string WriteParallel(
        DataTable table,
        CsvSaveOptions options,
        CsvWriteParallelOptions parallelOptions)
    {
        using DataTableReader reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteDataReaderParallel(writer, reader, options, parallelOptions);
        return writer.ToString();
    }

    private sealed class ThrowingFormattable : IFormattable
    {
        public string ToString(string? format, IFormatProvider? formatProvider) =>
            throw new InvalidOperationException("format failed");

        public override string ToString() => throw new InvalidOperationException("format failed");
    }

    private sealed class MutableFormattable : IFormattable
    {
        internal string Value { get; set; } = string.Empty;

        public string ToString(string? format, IFormatProvider? formatProvider) => Value;

        public override string ToString() => Value;
    }

    private sealed class CancelingFormattable : IFormattable
    {
        private readonly CancellationTokenSource _cancellation;

        internal CancelingFormattable(CancellationTokenSource cancellation) => _cancellation = cancellation;

        public string ToString(string? format, IFormatProvider? formatProvider)
        {
            _cancellation.Cancel();
            return "cancel";
        }

        public override string ToString() => ToString(null, CultureInfo.InvariantCulture);
    }

    private sealed class SignalingThrowingFormattable : IFormattable
    {
        private readonly ManualResetEventSlim _formattingFailed;

        internal SignalingThrowingFormattable(ManualResetEventSlim formattingFailed) =>
            _formattingFailed = formattingFailed;

        public string ToString(string? format, IFormatProvider? formatProvider)
        {
            try
            {
                throw new InvalidOperationException("format failed first");
            }
            finally
            {
                _formattingFailed.Set();
            }
        }

        public override string ToString() => ToString(null, CultureInfo.InvariantCulture);
    }

    private sealed class ThreadTrackingDataReader : IDataReader
    {
        private readonly IDataReader _inner;

        internal ThreadTrackingDataReader(IDataReader inner) => _inner = inner;

        internal HashSet<int> AccessThreadIds { get; } = new();

        private void Track() => AccessThreadIds.Add(Environment.CurrentManagedThreadId);

        public object this[int i] { get { Track(); return _inner[i]; } }
        public object this[string name] { get { Track(); return _inner[name]; } }
        public int Depth { get { Track(); return _inner.Depth; } }
        public bool IsClosed { get { Track(); return _inner.IsClosed; } }
        public int RecordsAffected { get { Track(); return _inner.RecordsAffected; } }
        public int FieldCount { get { Track(); return _inner.FieldCount; } }
        public void Close() { Track(); _inner.Close(); }
        public void Dispose() => _inner.Dispose();
        public bool GetBoolean(int i) { Track(); return _inner.GetBoolean(i); }
        public byte GetByte(int i) { Track(); return _inner.GetByte(i); }
        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) { Track(); return _inner.GetBytes(i, fieldOffset, buffer, bufferoffset, length); }
        public char GetChar(int i) { Track(); return _inner.GetChar(i); }
        public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) { Track(); return _inner.GetChars(i, fieldoffset, buffer, bufferoffset, length); }
        public IDataReader GetData(int i) { Track(); return _inner.GetData(i); }
        public string GetDataTypeName(int i) { Track(); return _inner.GetDataTypeName(i); }
        public DateTime GetDateTime(int i) { Track(); return _inner.GetDateTime(i); }
        public decimal GetDecimal(int i) { Track(); return _inner.GetDecimal(i); }
        public double GetDouble(int i) { Track(); return _inner.GetDouble(i); }
        public Type GetFieldType(int i) { Track(); return _inner.GetFieldType(i); }
        public float GetFloat(int i) { Track(); return _inner.GetFloat(i); }
        public Guid GetGuid(int i) { Track(); return _inner.GetGuid(i); }
        public short GetInt16(int i) { Track(); return _inner.GetInt16(i); }
        public int GetInt32(int i) { Track(); return _inner.GetInt32(i); }
        public long GetInt64(int i) { Track(); return _inner.GetInt64(i); }
        public string GetName(int i) { Track(); return _inner.GetName(i); }
        public int GetOrdinal(string name) { Track(); return _inner.GetOrdinal(name); }
        public DataTable? GetSchemaTable() { Track(); return _inner.GetSchemaTable(); }
        public string GetString(int i) { Track(); return _inner.GetString(i); }
        public object GetValue(int i) { Track(); return _inner.GetValue(i); }
        public int GetValues(object[] values) { Track(); return _inner.GetValues(values); }
        public bool IsDBNull(int i) { Track(); return _inner.IsDBNull(i); }
        public bool NextResult() { Track(); return _inner.NextResult(); }
        public bool Read() { Track(); return _inner.Read(); }
    }
}
