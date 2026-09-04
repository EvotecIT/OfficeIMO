using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

[JsonSourceGenerationOptions(
    WriteIndented = true,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    UseStringEnumConverter = true)]
[JsonSerializable(typeof(PdfRedactionRecipe))]
[JsonSerializable(typeof(PdfRedactionWorkflowRecord))]
internal sealed partial class PdfRedactionWorkflowJsonContext : JsonSerializerContext;
