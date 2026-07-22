using System.Text.Json.Serialization;

namespace OfficeIMO.Reader;

[JsonSourceGenerationOptions(
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    GenerationMode = JsonSourceGenerationMode.Metadata,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true)]
[JsonSerializable(typeof(OfficeDocumentReadResult))]
[JsonSerializable(typeof(ReaderChunkHierarchyResult))]
[JsonSerializable(typeof(OfficeDocumentStructuredExtractionResult))]
internal sealed partial class ReaderJsonSerializerContext : JsonSerializerContext {
}
