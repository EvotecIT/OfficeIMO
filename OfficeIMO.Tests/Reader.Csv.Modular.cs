using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderCsvModularTests {
    [Fact]
    public void DocumentReaderCsv_ReadCsvStream_ParsesCsvIntoStructuredChunks() {
        var csv =
            "Name,Role\n" +
            "Alice,Admin\n" +
            "Bob,Ops\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv), writable: false);

        var chunks = DocumentReaderCsvExtensions.ReadCsv(
            stream,
            sourceName: "users.csv",
            csvOptions: new CsvReadOptions {
                ChunkRows = 1,
                IncludeMarkdown = true
            }).ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Text, c.Kind));
        Assert.Contains(chunks, c => (c.Location.Path ?? string.Empty).Contains("users.csv", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(chunks, c => c.Tables != null && c.Tables.Count > 0 && c.Tables[0].Columns.Contains("Name", StringComparer.Ordinal));
    }

    [Fact]
    public void DocumentReaderCsv_ReadCsvStream_NonSeekable_EnforcesMaxInputBytes() {
        var csv =
            "Name,Role\n" +
            "Alice,Admin\n" +
            "Bob,Ops\n";
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(csv));

        var ex = Assert.Throws<IOException>(() => DocumentReaderCsvExtensions.ReadCsv(
            stream,
            sourceName: "users.csv",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 },
            csvOptions: new CsvReadOptions {
                ChunkRows = 1,
                IncludeMarkdown = true
            }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }
}
