using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Adf.Benchmarks;

public sealed class PlatformAdfDocument {
    [JsonPropertyName("version")]
    public int Version { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;

    [JsonPropertyName("content")]
    public List<PlatformAdfNode> Content { get; set; } = [];

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class PlatformAdfNode {
    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;

    [JsonPropertyName("text")]
    public string? Text { get; set; }

    [JsonPropertyName("attrs")]
    public Dictionary<string, JsonElement>? Attributes { get; set; }

    [JsonPropertyName("content")]
    public List<PlatformAdfNode>? Content { get; set; }

    [JsonPropertyName("marks")]
    public List<PlatformAdfMark>? Marks { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}

public sealed class PlatformAdfMark {
    [JsonPropertyName("type")]
    public string Type { get; set; } = string.Empty;

    [JsonPropertyName("attrs")]
    public Dictionary<string, JsonElement>? Attributes { get; set; }

    [JsonExtensionData]
    public Dictionary<string, JsonElement>? ExtensionData { get; set; }
}
