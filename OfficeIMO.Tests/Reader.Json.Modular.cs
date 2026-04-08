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

        var chunks = DocumentReaderJsonExtensions.ReadJson(
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

        var ex = Assert.Throws<IOException>(() => DocumentReaderJsonExtensions.ReadJson(
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
                DocumentReaderJsonExtensions.ReadJson(path, new ReaderOptions { ComputeHashes = true }),
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

        var chunks = DocumentReaderJsonExtensions.ReadJson(
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

    [Fact]
    public void DocumentReaderJson_ReadJsonStream_TrimsLogicalSourceName() {
        const string json = "{\"agent\":{\"name\":\"OfficeIMO\"}}";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(json), writable: false);

        var chunk = Assert.Single(DocumentReaderJsonExtensions.ReadJson(
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
