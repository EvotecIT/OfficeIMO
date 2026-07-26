using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;

namespace OfficeIMO.Tool.Agent;

internal static class AgentJson {
    internal static string Serialize<T>(T value) {
        JsonTypeInfo typeInfo = GetTypeInfo(value);
        return JsonSerializer.Serialize(value, typeInfo);
    }

    internal static JsonElement SerializeToElement<T>(T value) {
        JsonTypeInfo typeInfo = GetTypeInfo(value);
        return JsonSerializer.SerializeToElement(value, typeInfo);
    }

    internal static int Measure<T>(T value) => Serialize(value).Length;

    internal static JsonSerializerOptions CreateSerializerOptions() {
        var options = new JsonSerializerOptions(ModelContextProtocol.McpJsonUtilities.DefaultOptions);
        options.TypeInfoResolverChain.Insert(0, AgentJsonContext.Default);
        return options;
    }

    internal static string Limit(string? value, int maximumCharacters) {
        if (string.IsNullOrEmpty(value) || value!.Length <= maximumCharacters) {
            return value ?? string.Empty;
        }
        if (maximumCharacters <= 1) return value.Substring(0, maximumCharacters);
        return value.Substring(0, maximumCharacters - 1) + "…";
    }

    private static JsonTypeInfo GetTypeInfo<T>(T value) {
        Type type = value?.GetType() ?? typeof(T);
        return AgentJsonContext.Default.GetTypeInfo(type)
            ?? throw new NotSupportedException(
                "Compact agent JSON does not support type '" + type.FullName + "'.");
    }
}

[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(AgentInspectResult))]
[JsonSerializable(typeof(AgentSearchResult))]
[JsonSerializable(typeof(AgentFetchResult))]
[JsonSerializable(typeof(AgentCapabilitiesResult))]
[JsonSerializable(typeof(AgentConvertResult))]
internal sealed partial class AgentJsonContext : JsonSerializerContext;