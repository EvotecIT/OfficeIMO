#if NET8_0_OR_GREATER
using System;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvPendingStreamRowTests {
    [Theory]
    [InlineData("utf8-trim")]
    [InlineData("utf16-bom")]
    [InlineData("file")]
    [InlineData("buffer-boundary")]
    public void HeaderlessTextSourceRetainsPendingRowAcrossBufferGrowth(string route) {
        const int rowCount = 10_002;
        string firstName = route == "buffer-boundary" ? new string('x', 32_750) : "first";
        string text = CreateInput(firstName, rowCount);
        bool trim = route != "utf16-bom";
        Encoding encoding = trim ? new UTF8Encoding(false) : Encoding.Unicode;
        byte[] bytes = encoding.GetPreamble().Concat(encoding.GetBytes(text)).ToArray();
        var options = new CsvLoadOptions { HasHeaderRow = false, TrimWhitespace = trim };
        string? path = null;
        using var stream = new MemoryStream(bytes);
        try {
            if (route == "file") {
                path = Path.Combine(Path.GetTempPath(), $"officeimo-pending-row-{Guid.NewGuid():N}.csv");
                File.WriteAllBytes(path, bytes);
            }
            using DbDataReader reader = path is null
                ? CsvDocument.OpenDataReader(stream, options)
                : CsvDocument.OpenDataReader(path, options);
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("Column1", reader.GetName(0));
            Assert.True(reader.HasRows);
            Assert.Throws<InvalidOperationException>(() => reader.GetString(0));
            for (int index = 0; index < rowCount; index++) {
                Assert.True(reader.Read());
                string name = index == 0 ? firstName : index == 1 ? "next" : $"filler{index}";
                Assert.Equal(trim ? name : $" {name} ", reader.GetString(0));
                Assert.Equal(index + 1, reader.GetInt32(1));
                Assert.Equal(reader.GetString(0), reader.GetValue(0));
            }
            Assert.False(reader.Read());
            Assert.True(stream.CanRead);
        }
        finally {
            if (path is not null) File.Delete(path);
        }
    }

    [Fact]
    public void HeaderedTextStreamRetainsAllRecordsAfterBufferGrowth() {
        const int rowCount = 10_002;
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Name,Id\n" + CreateInput("first", rowCount)));
        using DbDataReader reader = CsvDocument.OpenDataReader(stream,
            new CsvLoadOptions { TrimWhitespace = true });
        Assert.Equal("Name", reader.GetName(0));
        for (int index = 0; index < rowCount; index++) {
            Assert.True(reader.Read());
            Assert.Equal(index == 0 ? "first" : index == 1 ? "next" : $"filler{index}", reader.GetString(0));
            Assert.Equal(index + 1, reader.GetInt32(1));
        }
        Assert.False(reader.Read());
    }

    private static string CreateInput(string firstName, int rowCount) {
        var text = new StringBuilder();
        for (int index = 0; index < rowCount; index++) {
            string name = index == 0 ? firstName : index == 1 ? "next" : $"filler{index}";
            text.Append(' ').Append(name).Append(" ,")
                .Append((index + 1).ToString(CultureInfo.InvariantCulture)).Append('\n');
        }
        return text.ToString();
    }
}
#endif
