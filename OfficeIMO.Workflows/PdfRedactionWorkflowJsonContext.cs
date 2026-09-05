using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

[JsonSourceGenerationOptions(
    WriteIndented = true,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    UseStringEnumConverter = true,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(PdfRedactionRecipe))]
[JsonSerializable(typeof(PdfRedactionDecisionManifest))]
[JsonSerializable(typeof(PdfRedactionWorkflowRecord))]
[JsonSerializable(typeof(PdfRedactionBatchRequest))]
[JsonSerializable(typeof(PdfRedactionBatchRecord))]
internal sealed partial class PdfRedactionWorkflowJsonContext : JsonSerializerContext;
