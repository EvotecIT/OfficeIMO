#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;
using AsyncInput = OfficeIMO.CSV.Tests.CsvIncrementalAsyncReaderTests.AsyncInput;

namespace OfficeIMO.CSV.Tests;

public class CsvRecordTextBatchTests
{
    [Theory]
    [InlineData(1)]
    [InlineData(4096)]
    [InlineData(32768)]
    public async Task CapturedFieldsAndPrimitiveValuesSurviveSourceAdvanceAndClose(int chunk)
    {
        const int rows = 259;
        var text = new StringBuilder("Id,Name,Money,Flag,When,Double,Empty,Nullable\n");
        var names = new string[rows];
        var dates = new DateTime[rows];
        for (int row = 0; row < rows; row++)
        {
            names[row] = row == 5 ? new string('x', 20000) : row % 47 == 0
                ? "first\r\nsecond \"quoted\" " + row : "Żółć-" + row + new string('y', 700);
            dates[row] = new DateTime(638456789012345678L + row, DateTimeKind.Utc);
            text.Append(9007199254740993L + row).Append(',').Append(Encode(names[row])).Append(',')
                .Append("1.25,").Append(row % 2 == 0 ? "true" : "0").Append(',')
                .Append(dates[row].ToString("O", CultureInfo.InvariantCulture)).Append(",1.5,");
            if (row % 3 != 2) text.Append(',').Append(row % 3 == 0 ? "NULL" : "");
            text.Append('\n');
        }
        var batches = new List<CsvRecordTextBatch>();
        var retained = new List<(string Value, string Expected)>();
        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text.ToString()), chunk);
        try
        {
            using (var reader = (CsvDataReader)await CsvDocument.OpenDataReaderAsync(input, new CsvLoadOptions
                { NullValue = "NULL", Culture = CultureInfo.InvariantCulture, DateTimeFormats = new[] { "O" } }))
            {
                bool reachedEnd = false;
                while (!reachedEnd)
                {
                    var capture = await reader.ReadRecordTextBatchAsync(193, CancellationToken.None);
                    reachedEnd = capture.ReachedEnd;
                    if (capture.Batch.Count == 0) capture.Batch.Dispose();
                    else batches.Add(capture.Batch);
                }
            }
            // Workers own their text even after subsequent buffer refills and source disposal.
            int index = 0;
            foreach (var batch in batches)
            {
                while (batch.Read())
                {
                    var record = new CsvRecord(batch);
                    Assert.Equal(8, record.FieldCount);
                    Assert.Equal(9007199254740993L + index, record.GetInt64(0));
                    Assert.Equal(names[index], record.GetSpan(1).ToString());
                    Assert.Equal(1.25m, record.GetDecimal(2));
                    Assert.Equal(index % 2 == 0, record.GetBoolean(3));
                    DateTime date = record.GetDateTime(4);
                    Assert.Equal(dates[index].Ticks, date.Ticks);
                    Assert.Equal(DateTimeKind.Utc, date.Kind);
                    Assert.Equal(1.5, record.GetDouble(5));
                    Assert.True(record.GetSpan(6).IsEmpty);
                    Assert.False(record.IsMissing(6));
                    Assert.False(record.IsNull(6));
                    Assert.Equal(index % 3 == 2, record.IsMissing(7));
                    Assert.Equal(index % 3 == 0, record.IsNull(7));
                    Assert.Equal(index % 3 == 0 ? "NULL" : "", record.GetSpan(7).ToString());
                    if (index % 31 == 0) retained.Add((record.GetString(1), names[index]));
                    index++;
                }
            }
            Assert.Equal(rows, index);
        }
        finally { foreach (var batch in batches) batch.Dispose(); }
        foreach (var value in retained) Assert.Equal(value.Expected, value.Value);
        Assert.True(input.AsyncReads > 0);
        Assert.True(input.CanRead);
    }

    private static string Encode(string value) => value.IndexOfAny(new[] { ',', '\r', '\n', '"' }) < 0
        ? value : "\"" + value.Replace("\"", "\"\"") + "\"";
}
#endif
