using OfficeIMO.Reader.Publisher;
using System.Text.Json;
using System.Text.Json.Nodes;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPublisherTransportTests {
    [Fact]
    public void NativePublisherResultUsesASchemaThatDeclaresItsKind() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "PublisherFixtures", "Simple.pub"));
        var reader = new OfficeDocumentReaderBuilder().AddPublisherHandler().Build();
        string json = reader.ReadDocument(bytes, "Simple.pub").ToJson();
        using var payload = JsonDocument.Parse(json);
        int version = payload.RootElement.GetProperty("schemaVersion").GetInt32();
        using var schema = JsonDocument.Parse(OfficeDocumentReadResultSchema.GetJsonSchema(version));
        Assert.Contains(payload.RootElement.GetProperty("kind").GetString(),
            schema.RootElement.GetProperty("properties").GetProperty("kind").GetProperty("enum")
                .EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(ReaderInputKind.Publisher, OfficeDocumentReadResultJson.Deserialize(json).Kind);
    }

    [Theory]
    [InlineData(5)]
    [InlineData(6)]
    [InlineData(7)]
    [InlineData(8)]
    [InlineData(9)]
    [InlineData(10)]
    public void EarlierTransportVersionsRejectPublisherAtEverySupportedDocumentLevel(int version) {
        var result = new OfficeDocumentReadResult { SchemaVersion = version, Kind = ReaderInputKind.Publisher };
        Assert.Throws<JsonException>(() => result.ToJson());
        result.Kind = ReaderInputKind.Text;
        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(
            result.ToJson().Replace("\"kind\":\"Text\"", "\"kind\":\"Publisher\"")));
        result.Chunks = new[] { new ReaderChunk { Kind = ReaderInputKind.Publisher } };
        Assert.Throws<JsonException>(() => result.ToJson());
        result.Chunks = new[] { new ReaderChunk { Kind = ReaderInputKind.Text } };
        JsonNode chunkEnvelope = JsonNode.Parse(result.ToJson())!;
        chunkEnvelope["chunks"]![0]!["kind"] = "Publisher";
        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(chunkEnvelope.ToJsonString()));
        if (version < 9) return;
        result.Chunks = Array.Empty<ReaderChunk>();
        result.NestedDocuments = new[] { new OfficeDocumentNestedResult {
            Path = "nested.pub", Document = new OfficeDocumentReadResult { Kind = ReaderInputKind.Publisher }
        } };
        Assert.Throws<JsonException>(() => result.ToJson());
        result.NestedDocuments[0].Document.Kind = ReaderInputKind.Text;
        JsonNode nestedEnvelope = JsonNode.Parse(result.ToJson())!;
        nestedEnvelope["nestedDocuments"]![0]!["document"]!["kind"] = "Publisher";
        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(nestedEnvelope.ToJsonString()));
    }
}
