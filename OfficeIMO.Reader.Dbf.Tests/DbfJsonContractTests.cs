using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeIMO.Reader.Dbf;

namespace OfficeIMO.Reader.Dbf.Tests {
    public sealed class DbfJsonContractTests {
        [Fact]
        public void NativeDbfResultUsesTheSchemaThatDeclaresItsKind() {
            var reader = new OfficeDocumentReaderBuilder().AddDbfHandler().Build();
            string json = reader.ReadDocumentJson(DbfReaderTests.Fixture("plain.dbf"));
            using var payload = JsonDocument.Parse(json);
            int version = payload.RootElement.GetProperty("schemaVersion").GetInt32();
            using var schema = JsonDocument.Parse(OfficeDocumentReadResultSchema.GetJsonSchema(version));
            Assert.Equal(11, version);
            Assert.Equal(version, schema.RootElement.GetProperty("properties").GetProperty("schemaVersion").GetProperty("const").GetInt32());
            Assert.Contains(payload.RootElement.GetProperty("kind").GetString(), schema.RootElement.GetProperty("properties").GetProperty("kind").GetProperty("enum").EnumerateArray().Select(value => value.GetString()));
            OfficeDocumentReadResult restored = OfficeDocumentReadResultJson.Deserialize(json);
            Assert.Equal(ReaderInputKind.Dbf, restored.Kind);
            Assert.All(restored.Chunks, chunk => Assert.Equal(ReaderInputKind.Dbf, chunk.Kind));
        }

        [Fact]
        public void OlderEnvelopesRejectDbfInRootsChunksAndNestedDocuments() {
            for (int version = 5; version <= 10; version++) {
                using var schema = JsonDocument.Parse(OfficeDocumentReadResultSchema.GetJsonSchema(version));
                Assert.DoesNotContain("Dbf", schema.RootElement.GetProperty("properties").GetProperty("kind").GetProperty("enum").EnumerateArray().Select(value => value.GetString()));
                var result = new OfficeDocumentReadResult { SchemaVersion = version, Kind = ReaderInputKind.Dbf };
                Assert.Throws<JsonException>(() => result.ToJson());
                result.Kind = ReaderInputKind.Text;
                string text = result.ToJson();
                Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(text.Replace("\"kind\":\"Text\"", "\"kind\":\"Dbf\"")));
                result.Chunks = new[] { new ReaderChunk { Kind = ReaderInputKind.Dbf } };
                Assert.Throws<JsonException>(() => result.ToJson());
                result.Chunks = new[] { new ReaderChunk { Kind = ReaderInputKind.Text } };
                text = result.ToJson();
                JsonNode chunkEnvelope = JsonNode.Parse(text)!;
                chunkEnvelope["chunks"]![0]!["kind"] = "Dbf";
                Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(chunkEnvelope.ToJsonString()));
                if (version < 9) continue;
                result.Chunks = Array.Empty<ReaderChunk>();
                result.NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "table.dbf", Document = new OfficeDocumentReadResult { Kind = ReaderInputKind.Dbf } } };
                Assert.Throws<JsonException>(() => result.ToJson());
                result.NestedDocuments[0].Document.Kind = ReaderInputKind.Text;
                text = result.ToJson();
                JsonNode nestedEnvelope = JsonNode.Parse(text)!;
                nestedEnvelope["nestedDocuments"]![0]!["document"]!["kind"] = "Dbf";
                Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(nestedEnvelope.ToJsonString()));
            }
        }
    }
}
