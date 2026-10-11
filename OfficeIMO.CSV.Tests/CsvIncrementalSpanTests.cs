#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;
using AsyncInput = OfficeIMO.CSV.Tests.CsvIncrementalAsyncReaderTests.AsyncInput;

namespace OfficeIMO.CSV.Tests;

public class CsvIncrementalSpanTests
{
    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(4096)]
    [InlineData(32768)]
    public async Task EveryFieldMatchesCanonicalParsingAcrossRefillsAndMixedRecords(int chunk)
    {
        const int columns = 19;
        var text = new StringBuilder();
        for (int i = 0; i < columns; i++) text.Append(i == 0 ? "Id" : "Column" + i).Append(i == columns - 1 ? '\n' : ',');
        for (int row = 0; row < 137; row++)
        {
            text.Append(9007199254740993L + row);
            for (int column = 1; column < columns; column++)
            {
                text.Append(',');
                if (column == 3 && row % 17 == 0) text.Append("\"one\r\ntwo \"\"quoted\"\"\"");
                else if (column == 4 && row == 3) text.Append(new string('x', 12000));
                else if (column != columns - 1) text.Append("\u2003Żółć-").Append(row).Append('-').Append(column).Append("\u00a0");
            }
            text.Append(row % 3 == 0 ? "\r\n" : row % 3 == 1 ? "\n" : "\r");
            if (row == 12) text.Append("\u2003\r\n");
        }
        // A final record without a separator takes the canonical EOF path.
        text.Append(9007199254741130L).Append(string.Concat(System.Linq.Enumerable.Repeat(",last", columns - 1)));
        var options = new CsvLoadOptions { TrimWhitespace = true, QuoteParsingMode = CsvQuoteParsingMode.Strict };
        using var expectedInput = new MemoryStream(Encoding.UTF8.GetBytes(text.ToString()));
        using var expected = CsvDocument.OpenDataReader(expectedInput, options);
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text.ToString()), chunk);
        using var actual = await CsvDocument.OpenDataReaderAsync(input, options);
        var retained = new List<(string Value, string Expected)>();
        int rows = 0;
        while (expected.Read())
        {
            Assert.True(await actual.ReadAsync());
            Assert.Equal(expected.GetInt64(0), actual.GetInt64(0));
            for (int column = 0; column < columns; column++)
            {
                string value = actual.GetString(column);
                string expectedValue = expected.GetString(column);
                Assert.Equal(expectedValue, value);
                if (rows % 31 == 0) retained.Add((value, expectedValue));
            }
            var expectedPosition = (ICsvDataReaderPositionMetadata)expected;
            var actualPosition = (ICsvDataReaderPositionMetadata)actual;
            Assert.Equal(expectedPosition.PhysicalLineNumber, actualPosition.PhysicalLineNumber);
            Assert.Equal(expectedPosition.PhysicalEndLineNumber, actualPosition.PhysicalEndLineNumber);
            rows++;
        }
        Assert.Equal(138, rows);
        Assert.False(await actual.ReadAsync());
        actual.Dispose();
        foreach (var value in retained) Assert.Equal(value.Expected, value.Value);
        Assert.True(input.AsyncReads > 0);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(false, 4096)]
    [InlineData(true, 4096)]
    [InlineData(false, 32768)]
    [InlineData(true, 32768)]
    public async Task CarriageReturnAtTransportBoundaryPreservesFollowingRecord(bool lineFeed, int chunk)
    {
        // The CR ends the first transport chunk, whether that refill is partial or full.
        string value = new string('a', chunk - 4);
        string csv = "V\n" + value + "x\r" + (lineFeed ? "\n" : "") + "next\r\nlast";
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(csv), chunk);
        using var reader = await CsvDocument.OpenDataReaderAsync(input);
        Assert.True(await reader.ReadAsync());
        Assert.Equal(value + "x", reader.GetString(0));
        Assert.True(await reader.ReadAsync());
        Assert.Equal("next", reader.GetString(0));
        Assert.Equal(3, ((ICsvDataReaderPositionMetadata)reader).PhysicalLineNumber);
        Assert.True(await reader.ReadAsync());
        Assert.Equal("last", reader.GetString(0));
        Assert.False(await reader.ReadAsync());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NormalizationAndInterningKeepTheirExistingFieldContracts(bool normalize)
    {
        const string csv = "V\n‘same’\n‘same’\n\"‘same’\"\n";
        var options = new CsvLoadOptions { NormalizeQuotes = normalize, InternStrings = !normalize };
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(csv), 4096);
        using var reader = await CsvDocument.OpenDataReaderAsync(input, options);
        string expected = normalize ? "'same'" : "‘same’";
        Assert.True(await reader.ReadAsync());
        string first = reader.GetString(0);
        Assert.Equal(expected, first);
        while (await reader.ReadAsync())
        {
            Assert.Equal(expected, reader.GetString(0));
            if (!normalize) Assert.Same(first, reader.GetString(0));
        }
        Assert.Equal(expected, first);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task CoveredAndIndependentReadTokensStillObserveEveryCancellationSource(int mode)
    {
        using var opening = new CancellationTokenSource();
        using var load = new CancellationTokenSource();
        using var operation = new CancellationTokenSource();
        using var input = new AsyncInput(Encoding.UTF8.GetBytes("V\nfirst\n"), 4096, blockAtEnd: true);
        using var reader = await CsvDocument.OpenDataReaderAsync(input,
            new CsvLoadOptions { CancellationToken = load.Token }, cancellationToken: opening.Token);
        Assert.True(await reader.ReadAsync());
        Assert.Equal("first", reader.GetString(0));
        CancellationToken readToken = mode == 0 ? opening.Token : mode == 1 ? load.Token : operation.Token;
        Task<bool> read = reader.ReadAsync(readToken);
        await input.Blocked.Task.WaitAsync(TimeSpan.FromSeconds(5));
        if (mode == 1) opening.Cancel();
        else if (mode == 2) operation.Cancel();
        else load.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => read.WaitAsync(TimeSpan.FromSeconds(5)));
        Assert.Equal(0, ((ICsvDataReaderPositionMetadata)reader).RecordNumber);
        Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
        await Assert.ThrowsAsync<InvalidOperationException>(() => reader.ReadAsync());
        Assert.True(input.CanRead);
    }
}
#endif
