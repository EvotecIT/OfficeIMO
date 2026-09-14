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
    /// <summary>Serializes a validated qualification manifest.</summary>
    public static string Serialize(HtmlRuntimeQualificationManifest manifest) {
        ArgumentNullException.ThrowIfNull(manifest);
        return JsonSerializer.Serialize(manifest, HtmlRuntimeQualificationJsonContext.Default.HtmlRuntimeQualificationManifest);
    }
    /// <summary>Serializes actual provider qualification counts.</summary>
    public static string Serialize(HtmlRuntimeQualificationResult result) {
        ArgumentNullException.ThrowIfNull(result);
        return JsonSerializer.Serialize(result, HtmlRuntimeQualificationJsonContext.Default.HtmlRuntimeQualificationResult);
    }
    /// <summary>Serializes actual consumer workflow qualification counts.</summary>
    public static string Serialize(HtmlRuntimeConsumerQualificationResult result) {
        ArgumentNullException.ThrowIfNull(result);
        return JsonSerializer.Serialize(result, HtmlRuntimeQualificationJsonContext.Default.HtmlRuntimeConsumerQualificationResult);
    }
    /// <summary>Serializes combined provider and consumer qualification evidence.</summary>
    public static string Serialize(HtmlRuntimeProfileQualificationResult result) {
        ArgumentNullException.ThrowIfNull(result);
        return JsonSerializer.Serialize(result, HtmlRuntimeQualificationJsonContext.Default.HtmlRuntimeProfileQualificationResult);
    }
}

[JsonSourceGenerationOptions(
    GenerationMode = JsonSourceGenerationMode.Serialization,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true)]
[JsonSerializable(typeof(HtmlRuntimeConformanceReport))]
internal sealed partial class HtmlRuntimeConformanceJsonContext : JsonSerializerContext;
