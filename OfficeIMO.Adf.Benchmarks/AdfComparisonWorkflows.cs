using System.Text.Json.Nodes;

namespace OfficeIMO.Adf.Benchmarks;

internal static class AdfComparisonWorkflows {
    internal static AdfDocument ParseOfficeIMO(string json) => AdfDocument.Parse(json);

    internal static JsonNode ParsePlatform(string json) =>
        JsonNode.Parse(json) ?? throw new InvalidDataException("Platform JSON parser returned null.");

    internal static string RoundTripOfficeIMO(string json) => AdfDocument.Parse(json).ToJson();

    internal static string RoundTripPlatform(string json) => ParsePlatform(json).ToJsonString();
}
