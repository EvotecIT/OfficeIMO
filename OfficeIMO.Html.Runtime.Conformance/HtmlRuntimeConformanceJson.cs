using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Html.Runtime.Conformance;

/// <summary>NativeAOT-safe JSON serialization for provider conformance reports.</summary>
public static class HtmlRuntimeConformanceJson {
    /// <summary>Serializes a conformance report with camel-case properties and string enums.</summary>
    public static string Serialize(HtmlRuntimeConformanceReport report) {
        ArgumentNullException.ThrowIfNull(report);
        return JsonSerializer.Serialize(report, HtmlRuntimeConformanceJsonContext.Default.HtmlRuntimeConformanceReport);
    }
}

[JsonSourceGenerationOptions(
    GenerationMode = JsonSourceGenerationMode.Serialization,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true)]
[JsonSerializable(typeof(HtmlRuntimeConformanceReport))]
internal sealed partial class HtmlRuntimeConformanceJsonContext : JsonSerializerContext;
