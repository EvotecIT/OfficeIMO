#if NET8_0_OR_GREATER
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvStreamParallelMappingTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitMappingConfiguresOnceOnTheEnumeratingThread(bool decodedText)
    {
        const string csv = "Id\n1\n2\n3\n4\n5\n";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(csv));
        using var text = new StringReader(csv);
        using var reader = decodedText ? CsvDocument.OpenReader(text) : CsvDocument.OpenReader(input);
        int caller = Environment.CurrentManagedThreadId;
        int calls = 0;
        int configurationThread = 0;
        var rows = reader.RowsAsParallel<MappedRow>(map => {
            Interlocked.Increment(ref calls);
            configurationThread = Environment.CurrentManagedThreadId;
            map.FromColumn<int>("Id", (row, value) => { row.Id = value; return row; });
        }, new ParallelRowMappingOptions { BatchSize = 2, MaxDegreeOfParallelism = 2 }).ToArray();
        Assert.Equal(new[] { 1, 2, 3, 4, 5 }, rows.Select(row => row.Id));
        Assert.Equal(1, calls);
        Assert.Equal(caller, configurationThread);
    }

    [Fact]
    public void InternedStringsRemainSharedAcrossParallelBatches()
    {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Name\nshared-value\nshared-value\nshared-value\n"));
        using var reader = CsvDocument.OpenReader(input, new CsvReaderOptions { InternStrings = true, DetectDelimiter = false });
        var rows = reader.RowsAsParallel<MappedRow>(new ParallelRowMappingOptions {
            BatchSize = 1, MaxDegreeOfParallelism = 2,
        }).ToArray();
        Assert.Equal(3, rows.Length);
        Assert.All(rows, row => Assert.Equal("shared-value", row.Name));
        Assert.Same(rows[0].Name, rows[1].Name);
        Assert.Same(rows[0].Name, rows[2].Name);
    }

    [Theory]
    [InlineData("1", false)]
    [InlineData("true", true)]
    public void StreamFactoriesKeepSnapshotBooleanConversions(string value, bool accepted)
    {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("Flag\n" + value + "\n"));
        using var reader = CsvDocument.OpenReader(input);
        var rows = reader.RowsAsParallel(record => record.GetBoolean(0),
            new ParallelRowMappingOptions { BatchSize = 1, MaxDegreeOfParallelism = 2 });
        if (accepted) Assert.True(Assert.Single(rows));
        else Assert.Throws<FormatException>(() => rows.ToArray());
    }

    public sealed class MappedRow
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }
}
#endif
