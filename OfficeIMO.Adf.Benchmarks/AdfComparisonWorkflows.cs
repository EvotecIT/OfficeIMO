using System.Text.Json.Nodes;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Adf.Benchmarks;

internal static class AdfComparisonWorkflows {
    private static readonly JsonSerializerOptions PlatformOptions = new() {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    internal static AdfDocument ParseOfficeIMO(string json) => AdfDocument.Parse(json);

    internal static JsonNode ParsePlatform(string json) =>
        JsonNode.Parse(json) ?? throw new InvalidDataException("Platform JSON parser returned null.");

    internal static PlatformAdfDocument ParsePlatformTyped(string json) =>
        JsonSerializer.Deserialize<PlatformAdfDocument>(json, PlatformOptions)
        ?? throw new InvalidDataException("Platform typed JSON parser returned null.");

    internal static string RoundTripOfficeIMO(string json) => AdfDocument.Parse(json).ToJson();

    internal static string RoundTripPlatform(string json) => ParsePlatform(json).ToJsonString();

    internal static string RoundTripPlatformTyped(string json) =>
        JsonSerializer.Serialize(ParsePlatformTyped(json), PlatformOptions);
}
