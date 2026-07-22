using System.Text.Json.Serialization;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    [JsonSourceGenerationOptions(
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNameCaseInsensitive = true)]
    [JsonSerializable(typeof(GoogleSlidesApiPresentationResponse))]
    [JsonSerializable(typeof(GoogleSlidesApiBatchResponse))]
    internal sealed partial class GoogleSlidesJsonSerializerContext : JsonSerializerContext {
    }
}
