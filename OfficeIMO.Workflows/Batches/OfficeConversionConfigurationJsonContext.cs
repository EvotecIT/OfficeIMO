using System.Text.Json.Serialization;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

// Source-generated native option metadata keeps checkpoint identity usable in trimmed and NativeAOT hosts.
internal sealed record OfficeConversionRenderingConfiguration(string Route, OfficeWorkflowOutputProfile OutputProfile,
    OfficeWorkflowConversionOptions Options, string Host);

[JsonSourceGenerationOptions(IncludeFields = true, GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(OfficeConversionRenderingConfiguration))]
internal partial class OfficeConversionConfigurationJsonContext : JsonSerializerContext { }

// Separate raw PDF metadata avoids recursively invoking the fingerprint converter.
[JsonSourceGenerationOptions(IncludeFields = true, GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(PdfOptions))]
internal partial class OfficeConversionPdfSettingsJsonContext : JsonSerializerContext { }
