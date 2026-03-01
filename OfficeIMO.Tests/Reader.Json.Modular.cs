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

            var chunks = DocumentReaderJsonExtensions.ReadJson(path).ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Json &&
                (c.Warnings?.Any(w => w.Contains("JSON parse error", StringComparison.OrdinalIgnoreCase)) ?? false));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
