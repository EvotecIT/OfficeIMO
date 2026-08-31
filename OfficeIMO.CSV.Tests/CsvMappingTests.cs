using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvMappingTests
{
#if NET8_0_OR_GREATER
    [Fact]
    public void Transient_Record_Projection_Uses_Bounded_Workers_And_Preserves_Order()
    {
        string text = "\"Id\",\"Name\"\n" + string.Join(
            "\n",
            Enumerable.Range(1, 256).Select(index => $"{index},\"Person {index}\"")) + "\n";
        using var firstWorkers = new Barrier(2);
        int calls = 0;
        int active = 0;
        int maximumActive = 0;

        PositionalPerson[] rows = CsvDocument.ReadTextRowsAsParallel<PositionalPerson>(
            text,
            header => {
                int id = header.GetOrdinal("Id");
                int name = header.GetOrdinal("Name");
                return record => {
                    int currentActive = Interlocked.Increment(ref active);
                    UpdateMaximum(ref maximumActive, currentActive);
                    try {
                        if (Interlocked.Increment(ref calls) <= 2) {
                            Assert.True(firstWorkers.SignalAndWait(TimeSpan.FromSeconds(10)));
                        }
                        return new PositionalPerson(record.GetInt32(id), record.GetString(name));
                    } finally {
                        Interlocked.Decrement(ref active);
                    }
                };
            },
            loadOptions: new CsvLoadOptions {
                // Progress reporting intentionally selects the scalar batch producer.
                ProgressCallback = _ => { }
            },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 64
            }).ToArray();

        Assert.Equal(Enumerable.Range(1, 256), rows.Select(row => row.Id));
        Assert.Equal(2, maximumActive);
    }

    [Fact]
    public void Transient_Record_Projection_Partitions_Quoted_Multiline_Rows_In_Order()
    {
        string text = "Id,Value\n" + string.Concat(
            Enumerable.Range(1, 32).Select(index =>
                $"{index},\"line {index}\ncontinued \"\"{index}\"\"\"\n"));

        PositionalPerson[] rows = CsvDocument.ReadTextRowsAsParallel<PositionalPerson>(
            text,
            header => {
                int id = header.GetOrdinal("Id");
                int value = header.GetOrdinal("Value");
                return record => new PositionalPerson(
                    record.GetInt32(id),
                    record.GetString(value));
            },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 4,
                BatchSize = 4
            }).ToArray();

        Assert.Equal(Enumerable.Range(1, 32), rows.Select(row => row.Id));
        Assert.Equal("line 17\ncontinued \"17\"", rows[16].Name);
    }

    [Fact]
    public void Transient_Record_Projection_Partition_Waves_Remain_Bounded_When_Enumeration_Stops()
    {
        string text = "Id\n" + string.Join("\n", Enumerable.Range(1, 20)) + "\n";
        int calls = 0;
        using IEnumerator<int> rows = CsvDocument.ReadTextRowsAsParallel<int>(
            text,
            _ => record => {
                Interlocked.Increment(ref calls);
                return record.GetInt32(0);
            },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 2
            }).GetEnumerator();

        Assert.True(rows.MoveNext());
        Assert.Equal(1, rows.Current);
        Assert.Equal(4, Volatile.Read(ref calls));
    }

    [Fact]
    public void Transient_Record_Projection_Falls_Back_After_The_Partition_Metadata_Budget()
    {
        const int rowCount = 16_385;
        string text = "Id\n" + string.Join("\n", Enumerable.Range(1, rowCount)) + "\n";

        int[] rows = CsvDocument.ReadTextRowsAsParallel<int>(
            text,
            _ => record => record.GetInt32(0),
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 1
            }).ToArray();

        Assert.Equal(rowCount, rows.Length);
        Assert.Equal(1, rows[0]);
        Assert.Equal(rowCount, rows[^1]);
    }

    [Fact]
    public void Transient_Record_Projection_Retains_Lenient_Bare_Quote_Fallback()
    {
        Assert.Throws<CsvParseException>(() =>
            CsvDocument.ReadTextRowsAsParallel<string>(
                "Value\na\"b\nc\n",
                _ => record => record.GetString(0),
                parallelOptions: new ParallelRowMappingOptions {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 1
                }).ToArray());
    }

    [Fact]
    public void Transient_Record_Projection_Observes_Cancellation_And_Conversion_Failures()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            CsvDocument.ReadTextRowsAsParallel<int>(
                "Id\n1\n2\n",
                _ => row => row.GetInt32(0),
                parallelOptions: new ParallelRowMappingOptions {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 1
                },
                cancellationToken: cancellation.Token).ToArray());

        FormatException exception = Assert.Throws<FormatException>(() =>
            CsvDocument.ReadTextRowsAsParallel<int>(
                "Id\n1\nnot-an-integer\n3\n",
                _ => row => row.GetInt32(0),
                parallelOptions: new ParallelRowMappingOptions {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 1
                }).ToArray());
        Assert.Contains("cannot be converted to Int32", exception.Message);
    }

    [Fact]
    public void TextBatchParserObservesTheMappingCancellationTokenIndependentlyOfLoadOptions()
    {
        var source = new CsvParser.CsvTextDataReaderRowSource(
            "Id\n1\n2\n",
            new CsvLoadOptions(),
            recordsToSkip: 1,
            sourceColumnCount: 1);
        using (source)
        using (var cancellation = new CancellationTokenSource())
        {
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() =>
                source.TryTakeParallelBatch(2, cancellation.Token, out _));
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void Transient_Record_Projection_Falls_Back_For_Fields_Too_Large_For_Packed_Metadata(
        int degreeOfParallelism)
    {
        string longValue = new('x', 70_000);
        string text = $"Id,Name\n1,\"{longValue}\"\n2,short\n";

        PositionalPerson[] rows = CsvDocument.ReadTextRowsAsParallel<PositionalPerson>(
            text,
            header => {
                int id = header.GetOrdinal("Id");
                int name = header.GetOrdinal("Name");
                return record => new PositionalPerson(record.GetInt32(id), record.GetString(name));
            },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = degreeOfParallelism,
                BatchSize = 1024
            }).ToArray();

        Assert.Equal(new[] { 1, 2 }, rows.Select(row => row.Id));
        Assert.Equal(longValue, rows[0].Name);
        Assert.Equal("short", rows[1].Name);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void Transient_Record_Projection_Preserves_Raw_Strings_When_Typed_Columns_Fall_Back(
        int degreeOfParallelism)
    {
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Done()
            .Build();

        string[] rows = CsvDocument.ReadTextRowsAsParallel<string>(
            "Id\n 42 \n",
            _ => record => record.GetString(0),
            loadOptions: new CsvLoadOptions { TrimWhitespace = true },
            readerOptions: new CsvDataReaderOptions { Schema = schema },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = degreeOfParallelism,
                BatchSize = 1
            }).ToArray();

        Assert.Equal(new[] { "42" }, rows);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void Transient_Record_Projection_Preserves_StringAccessAfterPreHeaderCommentFallback(
        int degreeOfParallelism)
    {
        CsvSchema schema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Done()
            .Build();

        string[] rows = CsvDocument.ReadTextRowsAsParallel<string>(
            "# generated export\nId\n001\n",
            _ => record => $"{record.GetString(0)}:{record.GetSpan(0).ToString()}",
            readerOptions: new CsvDataReaderOptions { Schema = schema },
            parallelOptions: new ParallelRowMappingOptions
            {
                MaxDegreeOfParallelism = degreeOfParallelism,
                BatchSize = 1
            }).ToArray();

        Assert.Equal(new[] { "001:001" }, rows);
    }

    [Theory]
    [InlineData("A,B\n1\n")]
    [InlineData("A,B\n1,2,3\n")]
    [InlineData("A,B\n1")]
    [InlineData("A,B\n1,2,3")]
    public void Transient_Record_Projection_Enforces_Strict_Column_Counts(string text)
    {
        Assert.Throws<CsvException>(() =>
            CsvDocument.ReadTextRowsAsParallel<string>(
                text,
                _ => row => row.GetString(0),
                loadOptions: new CsvLoadOptions {
                    ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict
                },
                parallelOptions: new ParallelRowMappingOptions {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 1024
                }).ToArray());
    }

    [Fact]
    public void Transient_Record_Projection_Enforces_Strict_Column_Counts_In_Vector_Body()
    {
        string text = "A,B\n" + new string('x', 40) + "\n";

        Assert.Throws<CsvException>(() =>
            CsvDocument.ReadTextRowsAsParallel<string>(
                text,
                _ => row => row.GetString(0),
                loadOptions: new CsvLoadOptions {
                    ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict
                },
                parallelOptions: new ParallelRowMappingOptions {
                    MaxDegreeOfParallelism = 2,
                    BatchSize = 1024
                }).ToArray());
    }

    [Fact]
    public void Strict_Text_DataReader_Yields_Valid_Prefix_Before_Mismatched_Row()
    {
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            "A,B\n1,2\n3\n",
            new CsvLoadOptions {
                ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict
            });

        Assert.True(reader.Read());
        Assert.Equal("1", reader.GetString(0));
        Assert.Equal("2", reader.GetString(1));
        Assert.Throws<CsvException>(() => reader.Read());
    }

    [Fact]
    public void Transient_Record_Projection_Applies_Configured_Size_To_First_Batch()
    {
        using var source = new CsvParser.CsvTextDataReaderRowSource(
            "1\n2\n",
            new CsvLoadOptions(),
            recordsToSkip: 0,
            sourceColumnCount: 1);

        Assert.True(source.TryTakeParallelBatch(
            1,
            CancellationToken.None,
            out ICsvDataReaderTextRowSource? rows));
        using (rows)
        {
            Assert.NotNull(rows);
            Assert.Equal(1, Assert.IsAssignableFrom<ICsvDataReaderParallelBatchInfo>(rows).RowCount);
        }
    }

    [Fact]
    public void Transient_Record_Projection_Observes_Cancellation_While_Yielding_Completed_Batch()
    {
        using var cancellation = new CancellationTokenSource();
        using IEnumerator<int> rows = CsvDocument.ReadTextRowsAsParallel<int>(
            "Id\n1\n2\n3\n",
            _ => row => row.GetInt32(0),
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 3
            },
            cancellationToken: cancellation.Token).GetEnumerator();

        Assert.True(rows.MoveNext());
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => rows.MoveNext());
    }

    [Fact]
    public void Transient_Record_Projection_Observes_Load_Cancellation_While_Yielding_Completed_Batch()
    {
        var text = new StringBuilder("Id\n");
        for (int id = 1; id <= 1024; id++) text.Append(id).Append('\n');
        using var cancellation = new CancellationTokenSource();
        using IEnumerator<int> rows = CsvDocument.ReadTextRowsAsParallel<int>(
            text.ToString(),
            _ => row => row.GetInt32(0),
            loadOptions: new CsvLoadOptions {
                CancellationToken = cancellation.Token
            },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 256
            }).GetEnumerator();

        Assert.True(rows.MoveNext());
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => rows.MoveNext());
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void Transient_Record_Projection_Does_Not_Read_Another_Row_After_Cancellation(int degreeOfParallelism)
    {
        var reports = new List<long>();
        using var cancellation = new CancellationTokenSource();
        using IEnumerator<int> rows = CsvDocument.ReadTextRowsAsParallel<int>(
            "Id\n1\n2\n3\n",
            _ => row => row.GetInt32(0),
            new CsvLoadOptions {
                ProgressReportInterval = 1,
                ProgressCallback = progress => reports.Add(progress.RecordsRead)
            },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = degreeOfParallelism,
                BatchSize = 1
            },
            cancellationToken: cancellation.Token).GetEnumerator();

        Assert.True(rows.MoveNext());
        Assert.Equal(1, rows.Current);
        long[] reportsBeforeCancellation = reports.ToArray();

        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => rows.MoveNext());
        Assert.Equal(reportsBeforeCancellation, reports);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void Transient_Record_Projection_Distinguishes_Missing_Null_And_Empty_Fields(int degreeOfParallelism)
    {
        string text = "Id,Value,Extra\n1,NULL\n2,,x\n3,\"NULL\",\"multi\nline\"\n";

        FieldState[] rows = CsvDocument.ReadTextRowsAsParallel<FieldState>(
            text,
            header => {
                int value = header.GetOrdinal("Value");
                int extra = header.GetOrdinal("Extra");
                return record => new FieldState(
                    record.IsMissing(value),
                    record.IsNull(value),
                    record.IsMissing(value) || record.IsNull(value) ? null : record.GetString(value),
                    record.IsMissing(extra),
                    record.IsNull(extra),
                    record.IsMissing(extra) || record.IsNull(extra) ? null : record.GetString(extra));
            },
            loadOptions: new CsvLoadOptions {
                NullValue = "NULL",
                TrimWhitespace = true
            },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = degreeOfParallelism,
                BatchSize = 1
            }).ToArray();

        Assert.Equal(new FieldState(false, true, null, true, false, null), rows[0]);
        Assert.Equal(new FieldState(false, false, string.Empty, false, false, "x"), rows[1]);
        Assert.Equal(new FieldState(false, true, null, false, false, "multi\nline"), rows[2]);
    }

    [Fact]
    public void Parallel_DataReader_Fallbacks_Actually_Map_Concurrently()
    {
        string text = "Id\n" + string.Join("\n", Enumerable.Range(1, 8)) + "\n";

        using (DbDataReader materialized = CsvDocument.Parse(text).CreateDataReader()) {
            AssertMapsConcurrently(materialized);
        }

        using (DbDataReader touched = CsvDocument.OpenTextDataReader(text)) {
            Assert.True(touched.HasRows);
            AssertMapsConcurrently(touched);
        }

        using (DbDataReader progress = CsvDocument.OpenTextDataReader(
                   text,
                   new CsvLoadOptions {
                       ProgressReportInterval = 1,
                       ProgressCallback = _ => { }
                   })) {
            AssertMapsConcurrently(progress);
        }

        string path = Path.Combine(Path.GetTempPath(), $"officeimo-parallel-{Guid.NewGuid():N}.csv");
        try {
            File.WriteAllText(path, text, new UTF8Encoding(false));
            using DbDataReader file = CsvDocument.OpenDataReader(path);
            AssertMapsConcurrently(file);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(text));
        using DbDataReader streamed = CsvDocument.OpenDataReader(stream);
        AssertMapsConcurrently(streamed);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void Parallel_Factory_Fallback_Uses_Independent_Ordered_Snapshots()
    {
        using DbDataReader reader = CsvDocument.Parse("Id\n1\n2\n3\n4\n").CreateDataReader();
        using var firstWorkers = new Barrier(2);
        int calls = 0;

        int[] rows = reader.RowsAsParallel(
            record => {
                Assert.NotSame(reader, record);
                Assert.Equal("Id", record.GetName(0));
                Assert.Equal(typeof(string), record.GetFieldType(0));
                if (Interlocked.Increment(ref calls) <= 2) {
                    Assert.True(firstWorkers.SignalAndWait(TimeSpan.FromSeconds(10)));
                }
                return record.GetInt32(0);
            },
            new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 1
            }).ToArray();

        Assert.Equal(new[] { 1, 2, 3, 4 }, rows);
    }

    [Fact]
    public void Parallel_Factory_Uses_The_Native_Sequential_Record_For_Unsafe_Object_Schemas()
    {
        var table = new DataTable();
        table.Columns.Add("Payload", typeof(object));
        using var payload = new MemoryStream(new byte[] { 1, 2, 3 });
        table.Rows.Add(payload);
        using DbDataReader reader = table.CreateDataReader();
        int callingThread = Environment.CurrentManagedThreadId;

        object[] rows = reader.RowsAsParallel(
            record => {
                Assert.Same(reader, record);
                Assert.Equal(callingThread, Environment.CurrentManagedThreadId);
                return record.GetValue(0);
            },
            new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 1
            }).ToArray();

        Assert.Same(payload, Assert.Single(rows));
    }

    [Fact]
    public void Parallel_Factory_Falls_Back_When_Optional_Type_Names_Are_Unimplemented()
    {
        using var reader = new ThrowingGetValuesDataReader(
            ["Id"],
            [[1], [2]],
            throwOnGetDataTypeName: true);

        int[] rows = reader.RowsAsParallel(
            record => record.GetInt32(0),
            new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 1
            }).ToArray();

        Assert.Equal(new[] { 1, 2 }, rows);
    }

    [Fact]
    public void Parallel_Automatic_Mapping_Falls_Back_When_Typed_Getters_Are_Unsupported()
    {
        using var reader = new ThrowingGetValuesDataReader(
            ["Id", "Name"],
            [[1, "Alpha"], [2, "Beta"]],
            throwOnTypedGetters: true);

        Person[] rows = reader.RowsAsParallel<Person>(
            new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 1
            }).ToArray();

        Assert.Equal(new[] { 1, 2 }, rows.Select(row => row.Id));
        Assert.Equal(new[] { "Alpha", "Beta" }, rows.Select(row => row.Name));
    }

    [Fact]
    public void Degree_One_Cancellation_Happens_Before_Invoking_The_Factory()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using DbDataReader reader = CsvDocument.Parse("Id\n1\n").CreateDataReader();
        int calls = 0;

        Assert.Throws<OperationCanceledException>(() => reader.RowsAsParallel(
            record => {
                Interlocked.Increment(ref calls);
                return record.GetInt32(0);
            },
            new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 1,
                BatchSize = 1
            },
            cancellation.Token).ToArray());

        Assert.Equal(0, calls);
    }

    [Fact]
    public void Parallel_DataReader_Yields_Queued_Prefix_Before_A_Source_Failure()
    {
        var sourceFailure = new InvalidOperationException("source failed");
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            "Id\n1\n2\n3\n4\n5\n",
            new CsvLoadOptions {
                ProgressReportInterval = 2,
                ProgressCallback = progress => {
                    if (progress.RecordsRead >= 4) throw sourceFailure;
                }
            });
        using IEnumerator<ProbeRow> rows = reader.RowsAsParallel<ProbeRow>(
            new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 2
            }).GetEnumerator();

        var yielded = new List<int>();
        InvalidOperationException actual = Assert.Throws<InvalidOperationException>(() => {
            while (rows.MoveNext()) yielded.Add(rows.Current.Id);
        });

        Assert.Same(sourceFailure, actual);
        Assert.Equal(new[] { 1, 2 }, yielded);
    }

    private static void AssertMapsConcurrently(DbDataReader reader)
    {
        using var firstWorkers = new Barrier(2);
        int calls = 0;
        ProbeRow[] rows = reader.RowsAsParallel<ProbeRow>(map => map
            .FromColumn<int>("Id", (row, value) => {
                if (Interlocked.Increment(ref calls) <= 2) {
                    Assert.True(firstWorkers.SignalAndWait(TimeSpan.FromSeconds(10)));
                }
                row.Id = value;
                return row;
            }),
            new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = 2,
                BatchSize = 1
            }).ToArray();

        Assert.Equal(Enumerable.Range(1, 8), rows.Select(row => row.Id));
    }

    private static void UpdateMaximum(ref int maximum, int candidate)
    {
        int observed;
        while (candidate > (observed = Volatile.Read(ref maximum)) &&
            Interlocked.CompareExchange(ref maximum, candidate, observed) != observed)
        {
        }
    }
#endif

    private sealed record Person
    {
        public int Id { get; init; }

        public string Name { get; init; } = string.Empty;

        public int Age { get; init; }

        public string City { get; init; } = string.Empty;
    }

    private sealed record EventRow
    {
        public DateTime Created { get; init; }
    }

    private sealed record PositionalPerson(int Id, string Name);

    private sealed record FieldState(
        bool ValueMissing,
        bool ValueNull,
        string? Value,
        bool ExtraMissing,
        bool ExtraNull,
        string? Extra);

    private sealed class ProbeRow
    {
        public int Id { get; set; }
    }

    [Fact]
    public void Maps_To_Typed_Record()
    {
        var doc = new CsvDocument()
            .WithHeader("Id", "Name", "Age", "City")
            .AddRow(1, "Przemek", 36, "Mikołów")
            .AddRow(2, "Dominika", 30, "Mikołów");

        var people = doc.RowsAs<Person>(map => map
            .FromColumn<int>("Id", (p, v) => p with { Id = v })
            .FromColumn<string>("Name", (p, v) => p with { Name = v })
            .FromColumn<int>("Age", (p, v) => p with { Age = v })
            .FromColumn<string>("City", (p, v) => p with { City = v })
        ).ToList();

        Assert.Equal(2, people.Count);
        Assert.Equal("Dominika", people[1].Name);
        Assert.Equal(30, people[1].Age);
    }

    [Fact]
    public void Map_Uses_Document_DateTime_Formats()
    {
        var doc = CsvDocument.Parse(
            "Created\n07-Jul-2026\n",
            new CsvLoadOptions { DateTimeFormats = new[] { "dd-MMM-yyyy" } });

        var row = Assert.Single(doc.RowsAs<EventRow>(map => map
            .FromColumn<DateTime>("Created", (item, value) => item with { Created = value })));

        Assert.Equal(new DateTime(2026, 7, 7), row.Created);
    }

    [Fact]
    public void Factory_Maps_Positional_Record_From_Document_And_DataReader()
    {
        var doc = new CsvDocument()
            .WithHeader("Id", "Name")
            .AddRow(42, "Ada");

        PositionalPerson fromDocument = Assert.Single(doc.RowsAs(factory: row =>
            new PositionalPerson(
                row.GetInt32(row.GetOrdinal("Id")),
                row.GetString(row.GetOrdinal("Name")))));

        using DbDataReader reader = doc.CreateDataReader();
        PositionalPerson fromReader = Assert.Single(reader.RowsAs(factory: row =>
            new PositionalPerson(
                row.GetInt32(row.GetOrdinal("Id")),
                row.GetString(row.GetOrdinal("Name")))));

        Assert.Equal(new PositionalPerson(42, "Ada"), fromDocument);
        Assert.Equal(fromDocument, fromReader);
    }

    [Fact]
    public void Parallel_DataReader_Mapping_Preserves_Order_For_All_Projection_Shapes()
    {
        string text = "Id,Name,Age,City\n" + string.Join(
            "\n",
            Enumerable.Range(1, 2049).Select(index => $"{index},Person {index},{20 + index % 50},City {index % 7}")) + "\n";
        var parallel = new ParallelRowMappingOptions {
            MaxDegreeOfParallelism = 4,
            BatchSize = 37
        };

        using DbDataReader automaticReader = CsvDocument.OpenTextDataReader(text);
        Person[] automatic = automaticReader.RowsAsParallel<Person>(parallel).ToArray();

        using DbDataReader explicitReader = CsvDocument.OpenTextDataReader(text);
        Person[] explicitRows = explicitReader.RowsAsParallel<Person>(map => map
            .FromColumn<int>("Id", (person, value) => person with { Id = value })
            .FromColumn<string>("Name", (person, value) => person with { Name = value })
            .FromColumn<int>("Age", (person, value) => person with { Age = value })
            .FromColumn<string>("City", (person, value) => person with { City = value }), parallel).ToArray();

        using DbDataReader factoryReader = CsvDocument.OpenTextDataReader(text);
        PositionalPerson[] factoryRows = factoryReader.RowsAsParallel(
            row => new PositionalPerson(row.GetInt32(0), row.GetString(1)),
            parallel).ToArray();

        Assert.Equal(2049, automatic.Length);
        Assert.Equal(Enumerable.Range(1, 2049), automatic.Select(person => person.Id));
        Assert.Equal(automatic, explicitRows);
        Assert.Equal(Enumerable.Range(1, 2049), factoryRows.Select(person => person.Id));
    }

    [Fact]
    public void Parallel_DataReader_Mapping_Observes_Cancellation()
    {
        string text = "Id,Name,Age,City\n1,Ada,36,London\n2,Grace,37,Arlington\n";
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using DbDataReader reader = CsvDocument.OpenTextDataReader(text);

        Assert.Throws<OperationCanceledException>(() => reader.RowsAsParallel<Person>(
            new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 },
            cancellation.Token).ToArray());
    }

    [Fact]
    public void Parallel_DataReader_Mapping_Observes_Cancellation_While_Yielding_Completed_Batch()
    {
        string text = "Id,Name,Age,City\n1,Ada,36,London\n2,Grace,37,Arlington\n3,Linus,40,Helsinki\n";
        using var cancellation = new CancellationTokenSource();
        using DbDataReader reader = CsvDocument.OpenTextDataReader(text);
        using IEnumerator<Person> rows = reader.RowsAsParallel<Person>(
            new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 3 },
            cancellation.Token).GetEnumerator();

        Assert.True(rows.MoveNext());
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => rows.MoveNext());
    }

    [Fact]
    public void Parallel_DataReader_Mapping_Rejects_Invalid_Bounds()
    {
        using DbDataReader reader = CsvDocument.OpenTextDataReader("Id,Name,Age,City\n1,Ada,36,London\n");
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.RowsAsParallel<Person>(
            new ParallelRowMappingOptions { MaxDegreeOfParallelism = 0 }).ToArray());
    }
}
