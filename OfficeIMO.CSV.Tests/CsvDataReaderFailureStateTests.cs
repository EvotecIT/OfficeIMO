using System;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvDataReaderFailureStateTests {
    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(2, true)]
    [InlineData(3, false)]
    [InlineData(3, true)]
    [InlineData(4, false)]
    [InlineData(4, true)]
    public async Task RejectedWidthHidesValuesAndPositionAndCannotResume(int source, bool asynchronous) {
        string csv = source == 1 ? "10,kept\n20\n30,later\n"
            : source == 2 ? "Id,Name\n\"10\",\"kept\"\n20\n30,later\n"
            : "Id,Name\n10,kept\n20\n30,later\n";
        var options = new CsvLoadOptions { ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict };
        if (source == 1) options.Header = new[] { "Id", "Name" };
        if (source == 3) options.ProgressCallback = _ => { };
        if (source == 4) options.InternStrings = true;
        using var input = source < 3 ? new MemoryStream(Encoding.UTF8.GetBytes(csv)) : null;
        using DbDataReader reader = source < 3
            ? CsvDocument.OpenDataReader(input!, options)
            : CsvDocument.OpenTextDataReader(csv, options);

        Assert.True(asynchronous ? await reader.ReadAsync() : reader.Read());
        Assert.Equal("10", reader.GetValue(0));
        Assert.Equal(10, reader.GetInt32(0));
        string retained = reader.GetString(1);
        Assert.Equal("kept", retained);
#if NET8_0_OR_GREATER
        {
            bool available = reader.TryGetUtf8Text(0, out var bytes);
            if (source < 2) Assert.True(available);
            if (available) Assert.Equal("10", Encoding.UTF8.GetString(bytes));
        }
#endif
        var position = (ICsvDataReaderPositionMetadata)reader;
        Assert.Equal(1, position.RecordNumber);

        if (asynchronous) await Assert.ThrowsAsync<CsvException>(() => reader.ReadAsync());
        else Assert.Throws<CsvException>(() => reader.Read());

        Assert.Equal(0, position.RecordNumber);
        Assert.Null(position.PhysicalLineNumber);
        Assert.Null(position.PhysicalEndLineNumber);
        Assert.Throws<InvalidOperationException>(() => reader.GetValue(0));
        Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
        Assert.Throws<InvalidOperationException>(() => reader.GetString(1));
        Assert.Throws<InvalidOperationException>(() => reader.GetInt32(0));
#if NET8_0_OR_GREATER
        Assert.Throws<InvalidOperationException>(() => reader.TryGetUtf8Text(0, out _));
#endif
        Assert.Throws<InvalidOperationException>(() => reader.Read());
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
        Assert.Equal("kept", retained);
        reader.Close();
        Assert.False(reader.Read());
        if (input is not null) Assert.True(input.CanRead);
    }
    [Fact]
    public async Task FailedAdvanceCannotBufferThroughHasRows() {
        using DbDataReader reader = CsvDocument.OpenTextDataReader("Id,Name\n10,kept\n20,later\n");
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            reader.ReadAsync(new CancellationToken(canceled: true)));

        Assert.Throws<InvalidOperationException>(() => reader.HasRows);
        Assert.Equal(0, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        Assert.Throws<InvalidOperationException>(() => reader.GetString(1));
    }

    [Fact]
    public async Task RejectedLookaheadIsTerminal() {
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            "Id,Name\n20\n10,kept\n",
            new CsvLoadOptions { ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict });
        Assert.Throws<CsvException>(() => reader.HasRows);

        Assert.Throws<InvalidOperationException>(() => reader.Read());
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
        Assert.Throws<InvalidOperationException>(() => reader.HasRows);
        Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
        Assert.Equal(0, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
    }

#if NET8_0_OR_GREATER
    [Fact]
    public async Task FailedAdvanceCannotResumeThroughParallelMapping() {
        const string csv = "Id,Name\n10,kept\n20,later\n";
        var options = new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2 };
        using (DbDataReader control = CsvDocument.OpenTextDataReader(csv)) {
            Assert.Equal(new[] { "kept", "later" }, control.RowsAsParallel(
                (IDataRecord row) => row.GetString(1), options).ToArray());
        }
        using DbDataReader reader = CsvDocument.OpenTextDataReader(csv);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            reader.ReadAsync(new CancellationToken(canceled: true)));

        string[]? returned = null;
        Exception? failure = Record.Exception(() => returned = reader.RowsAsParallel(
            (IDataRecord row) => row.GetString(1), options).ToArray());
        Assert.True(failure is InvalidOperationException,
            "Expected the failed reader to reject mapping. Returned rows: " +
            (returned is null ? "none" : string.Join(",", returned)) + "; failure: " + failure);
        Assert.Equal(0, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
    }
#endif

}
