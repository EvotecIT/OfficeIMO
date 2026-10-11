using OfficeIMO.Reader.Publisher;
using System.Text.Json;
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
}
