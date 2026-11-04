using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

// Source-generated native option metadata keeps checkpoint identity usable in trimmed and NativeAOT hosts.
internal sealed record OfficeConversionRenderingConfiguration(string Route, OfficeConversionProfile OutputProfile,
    OfficeWorkflowRenderingOptions Options, string Host, long MaximumXmlCharactersInPart);

[JsonSourceGenerationOptions(IncludeFields = true, GenerationMode = JsonSourceGenerationMode.Metadata, Converters = new[] { typeof(OfficeWorkflowCheckpointSettings.PdfConfigurationConverter), typeof(OfficeLengthFingerprintConverter), typeof(OfficeColorFingerprintConverter) })]
[JsonSerializable(typeof(OfficeConversionRenderingConfiguration))]
internal partial class OfficeConversionConfigurationJsonContext : JsonSerializerContext { }
