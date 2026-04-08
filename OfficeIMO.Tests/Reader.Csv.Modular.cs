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
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Csv, c.Kind));
        Assert.Contains(chunks, c => (c.Location.Path ?? string.Empty).Contains("users.csv", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(chunks, c => c.Tables != null && c.Tables.Count > 0 && c.Tables[0].Columns.Contains("Name", StringComparer.Ordinal));
        Assert.All(chunks, c => {
            Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
            Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
            Assert.Equal(stream.Length, c.SourceLengthBytes);
            Assert.Null(c.SourceLastWriteUtc);
        });
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

    [Fact]
    public void DocumentReaderCsv_ReadCsvStream_NormalizesBlankHeaders() {
        var csv =
            "Name,,Role\n" +
            "Alice,,Admin\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv), writable: false);

        var chunk = Assert.Single(DocumentReaderCsvExtensions.ReadCsv(
            stream,
            sourceName: "users.csv",
            csvOptions: new CsvReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = true
            }));

        var table = Assert.Single(chunk.Tables!);
        Assert.Equal(3, table.Columns.Count);
        Assert.Equal("Column2", table.Columns[1]);
        Assert.Contains("Column2", chunk.Markdown ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderCsv_ReadCsvStream_DeduplicatesHeaderNames() {
        var csv =
            "Name,Name ,Role\n" +
            "Alice,Alias,Admin\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv), writable: false);

        var chunk = Assert.Single(DocumentReaderCsvExtensions.ReadCsv(
            stream,
            sourceName: "users.csv",
            csvOptions: new CsvReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = true
            }));

        var table = Assert.Single(chunk.Tables!);
        Assert.Equal(new[] { "Name", "Name_2", "Role" }, table.Columns);
        Assert.Contains("Name | Name_2 | Role", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.Contains("Name_2", chunk.Markdown ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderCsv_ReadCsvStream_HeaderOnly_EmitsSchemaChunk() {
        const string csv = "Name,Role\n";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv), writable: false);

        var chunk = Assert.Single(DocumentReaderCsvExtensions.ReadCsv(
            stream,
            sourceName: "users.csv",
            csvOptions: new CsvReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = true
            }));

        var table = Assert.Single(chunk.Tables!);
        Assert.Equal(new[] { "Name", "Role" }, table.Columns);
        Assert.Empty(table.Rows);
        Assert.Equal(0, table.TotalRowCount);
        Assert.Contains("Name | Role", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        Assert.DoesNotContain("CSV content produced no rows", chunk.Text ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DocumentReaderCsv_ReadCsvStream_BlankFile_EmitsWarningChunk() {
        using var stream = new MemoryStream(Array.Empty<byte>(), writable: false);

        var chunk = Assert.Single(DocumentReaderCsvExtensions.ReadCsv(
            stream,
            sourceName: "users.csv",
            csvOptions: new CsvReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = true
            }));

        Assert.Equal(ReaderInputKind.Csv, chunk.Kind);
        Assert.Null(chunk.Tables);
        Assert.True((chunk.Warnings?.Any(w => w.Contains("produced no rows", StringComparison.OrdinalIgnoreCase)) ?? false));
        Assert.Contains("produced no rows", chunk.Text ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }
}
