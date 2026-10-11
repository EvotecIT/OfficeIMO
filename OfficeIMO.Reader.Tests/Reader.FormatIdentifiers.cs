using System.Text.Json;
using System.Text.Json.Nodes;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderFormatIdentifierTests {
    [Fact]
    public void PublishedInputKindNumbersRemainStableAndNewKindsAreDistinct() {
        string[] published = ("Unknown Word Excel PowerPoint Markdown Text Pdf Csv Json Xml Html Zip Epub Visio "
            + "Yaml Rtf OpenDocument AsciiDoc Latex Email OneNote Calendar VCard Opml DocBook IWork Xps Chm DjVu Dbf").Split(' ');
        for (int value = 0; value < published.Length; value++)
            Assert.Equal(value, (int)(ReaderInputKind)Enum.Parse(typeof(ReaderInputKind), published[value]));
        ReaderInputKind[] formats = { ReaderInputKind.Chm, ReaderInputKind.DjVu, ReaderInputKind.Dbf, ReaderInputKind.Publisher };
        Assert.Equal(new[] { 27, 28, 29, 30 }, formats.Select(kind => (int)kind));
        var builder = new OfficeDocumentReaderBuilder();
        foreach (ReaderInputKind kind in formats) builder.AddHandler(new ReaderHandlerRegistration {
            Id = "wire-test-" + kind, Kind = kind, Extensions = new[] { "." + kind.ToString().ToLowerInvariant() },
            ReadStream = (_, _, _, _) => new[] { new ReaderChunk { Kind = kind, Text = kind.ToString() } }
        });
        var reader = builder.Build();
        Assert.Equal(4, reader.GetCapabilities().Select(capability => capability.Kind).Distinct().Count());
        foreach (ReaderInputKind kind in formats) {
            OfficeDocumentReadResult result = reader.ReadDocument(new byte[] { 65 }, "proof." + kind.ToString().ToLowerInvariant());
            Assert.Equal(kind, result.Kind);
            Assert.Equal(kind.ToString(), Assert.Single(result.Chunks).Text);
        }
    }

    [Theory]
    [InlineData(ReaderInputKind.Chm, "Chm")]
    [InlineData(ReaderInputKind.DjVu, "DjVu")]
    [InlineData(ReaderInputKind.Dbf, "Dbf")]
    [InlineData(ReaderInputKind.Publisher, "Publisher")]
    public void CurrentSchemaDeclaresAndRoundTripsEachKindAtEveryDocumentLevel(ReaderInputKind kind, string name) {
        var source = new OfficeDocumentReadResult { Kind = kind, Chunks = new[] { new ReaderChunk { Kind = kind } },
            NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "nested", Document = new() { Kind = kind } } } };
        string json = source.ToJson();
        using var payload = JsonDocument.Parse(json);
        Assert.Equal(14, payload.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(name, payload.RootElement.GetProperty("kind").GetString());
        using var schema = JsonDocument.Parse(OfficeDocumentReadResultSchema.GetJsonSchema(14));
        Assert.Contains(name, schema.RootElement.GetProperty("properties").GetProperty("kind").GetProperty("enum")
            .EnumerateArray().Select(item => item.GetString()));
        OfficeDocumentReadResult restored = OfficeDocumentReadResultJson.Deserialize(json);
        Assert.Equal(kind, restored.Kind);
        Assert.Equal(kind, Assert.Single(restored.Chunks).Kind);
        Assert.Equal(kind, Assert.Single(restored.NestedDocuments).Document.Kind);
    }

    [Theory]
    [InlineData(ReaderInputKind.Chm, 11)]
    [InlineData(ReaderInputKind.DjVu, 12)]
    [InlineData(ReaderInputKind.Dbf, 13)]
    [InlineData(ReaderInputKind.Publisher, 14)]
    public void EarlierSchemasRejectNewKindsInRootsChunksAndNestedDocuments(ReaderInputKind kind, int minimumVersion) {
        string name = kind.ToString();
        for (int version = 5; version < minimumVersion; version++) {
            using var schema = JsonDocument.Parse(OfficeDocumentReadResultSchema.GetJsonSchema(version));
            Assert.DoesNotContain(name, schema.RootElement.GetProperty("properties").GetProperty("kind").GetProperty("enum")
                .EnumerateArray().Select(item => item.GetString()));
            var source = new OfficeDocumentReadResult { SchemaVersion = version, Kind = kind };
            Assert.Throws<JsonException>(() => source.ToJson());
            source.Kind = ReaderInputKind.Text;
            Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(
                source.ToJson().Replace("\"kind\":\"Text\"", "\"kind\":\"" + name + "\"")));
            source.Chunks = new[] { new ReaderChunk { Kind = kind } };
            Assert.Throws<JsonException>(() => source.ToJson());
            source.Chunks = new[] { new ReaderChunk { Kind = ReaderInputKind.Text } };
            JsonNode chunkEnvelope = JsonNode.Parse(source.ToJson())!;
            chunkEnvelope["chunks"]![0]!["kind"] = name;
            Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(chunkEnvelope.ToJsonString()));
            if (version < 9) continue;
            source.Chunks = Array.Empty<ReaderChunk>();
            source.NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "nested", Document = new() { Kind = kind } } };
            Assert.Throws<JsonException>(() => source.ToJson());
            source.NestedDocuments[0].Document.Kind = ReaderInputKind.Text;
            JsonNode nestedEnvelope = JsonNode.Parse(source.ToJson())!;
            nestedEnvelope["nestedDocuments"]![0]!["document"]!["kind"] = name;
            Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(nestedEnvelope.ToJsonString()));
        }
    }
}
