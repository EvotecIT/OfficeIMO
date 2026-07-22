using System.Text.Json.Serialization;

namespace OfficeIMO.Confluence;

[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    PropertyNameCaseInsensitive = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(ConfluencePage))]
[JsonSerializable(typeof(ConfluenceCollectionResponse<ConfluencePage>))]
[JsonSerializable(typeof(ConfluenceCollectionResponse<ConfluenceAttachment>))]
[JsonSerializable(typeof(ConfluencePageCreatePayload))]
[JsonSerializable(typeof(ConfluencePageUpdatePayload))]
[JsonSerializable(typeof(ConfluenceV1AttachmentResponse))]
internal sealed partial class ConfluenceJsonSerializerContext : JsonSerializerContext {
}
