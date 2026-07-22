using System.Text.Json.Serialization;

namespace OfficeIMO.Word.GoogleDocs {
    [JsonSourceGenerationOptions(
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNameCaseInsensitive = true)]
    [JsonSerializable(typeof(GoogleDocsApiCreateDocumentPayload))]
    [JsonSerializable(typeof(GoogleDocsApiBatchUpdatePayload))]
    [JsonSerializable(typeof(GoogleDocsApiCreateDocumentResponse))]
    [JsonSerializable(typeof(GoogleDocsApiDocumentResponse))]
    [JsonSerializable(typeof(GoogleDocsApiBatchUpdateResponse))]
    internal sealed partial class GoogleDocsJsonSerializerContext : JsonSerializerContext {
    }
}
