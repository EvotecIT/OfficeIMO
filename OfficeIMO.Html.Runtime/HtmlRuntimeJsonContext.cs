using System.Text.Json.Serialization;

namespace OfficeIMO.Html.Runtime;

[JsonSourceGenerationOptions(GenerationMode = JsonSourceGenerationMode.Default)]
[JsonSerializable(typeof(HtmlRuntimeCommand))]
[JsonSerializable(typeof(HtmlRuntimeResponse))]
internal sealed partial class HtmlRuntimeProtocolJsonContext : JsonSerializerContext;

[JsonSourceGenerationOptions(
    GenerationMode = JsonSourceGenerationMode.Default,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(HtmlPageObservationRequest))]
[JsonSerializable(typeof(HtmlAutomationToolArguments))]
[JsonSerializable(typeof(HtmlNavigationToolArguments))]
[JsonSerializable(typeof(HtmlCaptureToolArguments))]
internal sealed partial class HtmlAutomationToolJsonContext : JsonSerializerContext;

[JsonSourceGenerationOptions(
    GenerationMode = JsonSourceGenerationMode.Serialization,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true)]
[JsonSerializable(typeof(HtmlRuntimeProviderDescriptor))]
[JsonSerializable(typeof(HtmlPageObservation))]
[JsonSerializable(typeof(HtmlAutomationResult))]
[JsonSerializable(typeof(HtmlRuntimeTrace))]
[JsonSerializable(typeof(HtmlRuntimeArtifactManifest))]
[JsonSerializable(typeof(HtmlAutomationToolResultPayload))]
[JsonSerializable(typeof(HtmlAutomationRunResultPayload))]
internal sealed partial class HtmlRuntimePublicJsonContext : JsonSerializerContext;
