using OfficeIMO.Reader;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderDjVuContractTests {
    [Fact]
    public void DjVuSchemaPreservesTheExistingChmTransportIdentity() {
        Assert.Equal(27, (int)ReaderInputKind.Chm);
        Assert.Equal(28, (int)ReaderInputKind.DjVu);
        using var previous = JsonDocument.Parse(OfficeDocumentReadResultSchema.GetJsonSchema(11));
        string?[] previousKinds = previous.RootElement.GetProperty("properties").GetProperty("kind")
            .GetProperty("enum").EnumerateArray().Select(value => value.GetString()).ToArray();
        Assert.Contains("Chm", previousKinds);
        Assert.DoesNotContain("DjVu", previousKinds);
        using var current = JsonDocument.Parse(OfficeDocumentReadResultSchema.GetJsonSchema(12));
        string?[] currentKinds = current.RootElement.GetProperty("properties").GetProperty("kind")
            .GetProperty("enum").EnumerateArray().Select(value => value.GetString()).ToArray();
        Assert.Contains("Chm", currentKinds);
        Assert.Contains("DjVu", currentKinds);

        var existing = new OfficeDocumentReadResult { Kind = ReaderInputKind.Chm, SchemaVersion = 11 };
        string existingJson = OfficeDocumentReadResultJson.Serialize(existing);
        Assert.Equal(ReaderInputKind.Chm, OfficeDocumentReadResultJson.Deserialize(existingJson).Kind);
        using var serialized = JsonDocument.Parse(existingJson);
        Assert.Equal(11, serialized.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("Chm", serialized.RootElement.GetProperty("kind").GetString());
    }

    [Fact]
    public void DjVuTransportRejectsVersionElevenAtEnvelopeChunkAndNestedBoundaries() {
        var native = new OfficeDocumentReadResult { Kind = ReaderInputKind.DjVu };
        string json = OfficeDocumentReadResultJson.Serialize(native);
        Assert.Equal(ReaderInputKind.DjVu, OfficeDocumentReadResultJson.Deserialize(json).Kind);
        var envelope = JsonNode.Parse(json)!;
        envelope["schemaVersion"] = 11;
        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(envelope.ToJsonString()));
        native.SchemaVersion = 11;
        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Serialize(native));

        var projected = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Text,
            Chunks = new[] { new ReaderChunk { Kind = ReaderInputKind.DjVu, Text = "stored text" } }
        };
        string projectedJson = OfficeDocumentReadResultJson.Serialize(projected);
        var projectedEnvelope = JsonNode.Parse(projectedJson)!;
        projectedEnvelope["schemaVersion"] = 11;
        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(projectedEnvelope.ToJsonString()));
        projected.SchemaVersion = 11;
        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Serialize(projected));

        var archive = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Zip,
            SchemaVersion = 11,
            NestedDocuments = new[] {
                new OfficeDocumentNestedResult {
                    Path = "book.djvu",
                    Document = new OfficeDocumentReadResult { Kind = ReaderInputKind.DjVu }
                }
            }
        };
        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Serialize(archive));
    }
}
