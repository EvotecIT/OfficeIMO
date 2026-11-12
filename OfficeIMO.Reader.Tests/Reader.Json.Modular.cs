using OfficeIMO.Reader;
using OfficeIMO.Reader.Json;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderJsonModularTests {
    [Fact]
    public void DocumentReaderJson_ReadJsonStream_ParsesJsonIntoStructuredChunks() {
        var json =
            "{\n" +
            "  \"agent\": {\n" +
            "    \"name\": \"OfficeIMO\",\n" +
            "    \"version\": 1\n" +
            "  }\n" +
            "}";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(json), writable: false);

        var chunks = JsonReaderAdapter.Read(
            stream,
            sourceName: "agent.json",
            jsonOptions: new JsonReadOptions {
                ChunkRows = 2,
                IncludeMarkdown = true
            }).ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Json, c.Kind));
        Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("$.agent.name", StringComparison.Ordinal));
        Assert.Contains(chunks, c => c.Tables != null && c.Tables.Count > 0 && c.Tables[0].Columns.Contains("Path", StringComparer.Ordinal));
        var table = Assert.Single(chunks.SelectMany(c => c.Tables ?? Array.Empty<ReaderTable>()));
        Assert.Equal(3, table.ColumnProfiles.Count);
        Assert.Equal(ReaderTableColumnKind.Text, table.ColumnProfiles[0].Kind);
        Assert.Equal(ReaderTableColumnKind.Text, table.ColumnProfiles[1].Kind);
        Assert.Equal(ReaderTableColumnKind.Mixed, table.ColumnProfiles[2].Kind);
        Assert.Equal(2, table.ColumnProfiles[2].NonEmptyCellCount);
        Assert.Equal(1, table.ColumnProfiles[2].NumericCellCount);
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
    public void DocumentReaderJson_ReadJsonStream_NonSeekable_EnforcesMaxInputBytes() {
        var json =
            "{\n" +
            "  \"agent\": {\n" +
            "    \"name\": \"OfficeIMO\"\n" +
            "  }\n" +
            "}";
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(json));

        var ex = Assert.Throws<IOException>(() => JsonReaderAdapter.Read(
            stream,
            sourceName: "agent.json",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 },
            jsonOptions: new JsonReadOptions {
                ChunkRows = 2,
                IncludeMarkdown = true
            }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderJson_ReadJson_EmitsWarningForMalformedJson() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-bad-json-" + Guid.NewGuid().ToString("N") + ".json");
        try {
            File.WriteAllText(path, "{ bad json ");

            var warningChunk = Assert.Single(
                JsonReaderAdapter.Read(path, new ReaderOptions { ComputeHashes = true }),
                c => c.Kind == ReaderInputKind.Json &&
                     (c.Warnings?.Any(w => w.Contains("JSON parse error", StringComparison.OrdinalIgnoreCase)) ?? false));

            Assert.False(string.IsNullOrWhiteSpace(warningChunk.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(warningChunk.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(warningChunk.ChunkHash));
            Assert.True(warningChunk.TokenEstimate.HasValue && warningChunk.TokenEstimate.Value >= 1);
            Assert.True(warningChunk.SourceLengthBytes.HasValue && warningChunk.SourceLengthBytes.Value > 0);
            Assert.True(warningChunk.SourceLastWriteUtc.HasValue);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderJson_ReadJsonStream_EscapesComplexPropertyNamesInPaths() {
        const string json = "{\"service.name\":{\"port[0]\":443}}";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(json), writable: false);

        var chunks = JsonReaderAdapter.Read(
            stream,
            sourceName: "agent.json",
            jsonOptions: new JsonReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = true
            }).ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("$[\"service.name\"][\"port[0]\"]", StringComparison.Ordinal));
        Assert.Contains(chunks, c => (c.Markdown ?? string.Empty).Contains("$[\"service.name\"][\"port[0]\"]", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("MaxDepth", 2, "{\"a\":{\"b\":1}}", false)]
    [InlineData("MaxDepth", 2, "{\"a\":{\"b\":{}}}", true)]
    [InlineData("MaxDepth", 2, "[[[[[[[[[[1]]]]]]]]]]", true)]
    [InlineData("MaxNodes", 4, "[1,true,null]", false)]
    [InlineData("MaxNodes", 4, "[1,true,null,\"x\"]", true)]
    [InlineData("MaxScalarLength", 2, "{\"k\":\"\\u0041\\u0042\"}", false)]
    [InlineData("MaxScalarLength", 2, "{\"k\":\"abc\"}", true)]
    [InlineData("MaxScalarLength", 2, "{\"abc\":1}", true)]
    [InlineData("MaxScalarLength", 2, "[123]", true)]
    public void JsonLimitsThrowTheNamedLimitOnceExceeded(string limit, int value, string json, bool exceeds) {
        var options = new JsonReadOptions {
            MaxDepth = limit == "MaxDepth" ? value : 32,
            MaxNodes = limit == "MaxNodes" ? value : 200_000,
            MaxScalarLength = limit == "MaxScalarLength" ? value : 1_048_576
        };
        var reader = new OfficeDocumentReaderBuilder().AddJsonHandler(options).Build();
        Func<OfficeDocumentReadResult> read = () => reader.ReadDocument(Encoding.UTF8.GetBytes(json), "bounded.json", new ReaderOptions { ComputeHashes = false });

        if (!exceeds) {
            OfficeDocumentReadResult result = read();
            Assert.Empty(result.Diagnostics);
            Assert.NotEmpty(result.Tables);
            return;
        }

        ReaderResourceLimitException error = Assert.Throws<ReaderResourceLimitException>(() => read());
        Assert.Equal(limit, error.LimitName);
        Assert.Equal(value, error.Maximum);
    }

    [Fact]
    public void JsonReadStopsWhenCancelled() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("{\"a\":[1,2,3]}"), writable: false);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() => JsonReaderAdapter.Read(stream, "a.json", cancellationToken: cancellation.Token).ToList());
    }

    [Fact]
    public void DocumentReaderJson_ReadJsonStream_TrimsLogicalSourceName() {
        const string json = "{\"agent\":{\"name\":\"OfficeIMO\"}}";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(json), writable: false);

        var chunk = Assert.Single(JsonReaderAdapter.Read(
            stream,
            sourceName: " agent.json ",
            jsonOptions: new JsonReadOptions {
                ChunkRows = 10,
                IncludeMarkdown = true
            }));

        Assert.Equal("agent.json", chunk.Location.Path);
        Assert.Contains("$.agent.name", chunk.Text ?? string.Empty, StringComparison.Ordinal);
    }
}
