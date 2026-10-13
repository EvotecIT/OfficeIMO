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
    [InlineData(1, 193)]
    [InlineData(4096, 193)]
    [InlineData(32768, 193)]
    [InlineData(32768, 1024)]
    public async Task CapturedFieldsAndPrimitiveValuesSurviveSourceAdvanceAndClose(int chunk, int batchSize)
    {
        int rows = batchSize == 1024 ? 2053 : 259;
        var text = new StringBuilder("Id,Name,Money,Flag,When,Double,Empty,Nullable\n");
        var names = new string[rows];
        var dates = new DateTime[rows];
        for (int row = 0; row < rows; row++)
        {
            names[row] = row == 5 ? new string('x', batchSize == 1024 ? 300_000 : 20_000) : row % 47 == 0
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
                    var batch = reader.CreateRecordTextBatch(batchSize);
                    try { reachedEnd = await reader.ReadRecordTextBatchAsync(batch, CancellationToken.None); }
                    catch { batch.Dispose(); throw; }
                    if (batch.Count == 0) batch.Dispose();
                    else batches.Add(batch);
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

    [Fact]
    public async Task LateRawFieldMaterializationKeepsOtherFieldsCorrectAcrossBatchReuse()
    {
        var text = new StringBuilder("Id,Name,Tail,Empty,Nullable,Missing\n");
        for (int id = 1; id <= 6; id++)
            text.Append(id).Append(",name").Append(id).Append(",tail").Append(id)
                .Append(",,").Append(id % 2 == 0 ? "NULL" : "").Append('\n');

        using var input = new AsyncInput(Encoding.UTF8.GetBytes(text.ToString()), 4096);
        using var reader = (CsvDataReader)await CsvDocument.OpenDataReaderAsync(input, new CsvLoadOptions
        {
            NullValue = "NULL",
            ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.PadMissingFieldsAndIgnoreExtraFields,
        });
        var retained = new List<string>();
        using (var batch = reader.CreateRecordTextBatch(3))
        {
            for (int pass = 0; pass < 2; pass++)
            {
                await reader.ReadRecordTextBatchAsync(batch, CancellationToken.None);
                Assert.Equal(3, batch.Count);
                for (int offset = 1; offset <= 3; offset++)
                {
                    Assert.True(batch.Read());
                    var record = new CsvRecord(batch);
                    int id = pass * 3 + offset;
                    Assert.Equal(id, record.GetInt32(0));
                    Assert.Equal("name" + id, record.GetSpan(1).ToString());
                    Assert.Equal(string.Empty, record.GetString(3));
                    Assert.Equal(string.Empty, record.GetString(5));
                    if (id >= 2) retained.Add(record.GetString(1));
                    // The first nonempty string request occurs only after a span-only row.
                    Assert.Equal("tail" + id, record.GetSpan(2).ToString());
                    if (id >= 2) Assert.Equal("tail" + id, record.GetString(2));
                    Assert.Equal(id % 2 == 0 ? "NULL" : "", record.GetString(4));
                    Assert.Equal(id % 2 == 0, record.IsNull(4));
                    Assert.True(record.IsMissing(5));
                    Assert.False(record.IsMissing(3));
                    Assert.False(record.IsNull(3));
                    Assert.Equal("name" + id, record.GetSpan(1).ToString());
                }
                Assert.False(batch.Read());
            }
        }
        Assert.Equal(new[] { "name2", "name3", "name4", "name5", "name6" }, retained);
    }

    private static string Encode(string value) => value.IndexOfAny(new[] { ',', '\r', '\n', '"' }) < 0
        ? value : "\"" + value.Replace("\"", "\"\"") + "\"";
}
#endif
