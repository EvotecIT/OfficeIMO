using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;

namespace OfficeIMO.GoogleWorkspace.Drive {
    [JsonSourceGenerationOptions(
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        GenerationMode = JsonSourceGenerationMode.Metadata)]
    [JsonSerializable(typeof(object))]
    [JsonSerializable(typeof(GoogleDriveAboutFormats))]
    [JsonSerializable(typeof(GoogleDriveFile))]
    [JsonSerializable(typeof(GoogleDriveFileList))]
    [JsonSerializable(typeof(GoogleSharedDrive))]
    [JsonSerializable(typeof(GoogleDrivePermission))]
    [JsonSerializable(typeof(GoogleDrivePermissionList))]
    [JsonSerializable(typeof(GoogleDriveComment))]
    [JsonSerializable(typeof(GoogleDriveCommentList))]
    [JsonSerializable(typeof(GoogleDriveReply))]
    [JsonSerializable(typeof(GoogleDriveRevisionList))]
    [JsonSerializable(typeof(GoogleDriveStartPageToken))]
    [JsonSerializable(typeof(GoogleDriveChangeList))]
    [JsonSerializable(typeof(GoogleDriveFilePayload))]
    [JsonSerializable(typeof(GoogleDrivePermissionPayload))]
    [JsonSerializable(typeof(GoogleDriveCommentPayload))]
    [JsonSerializable(typeof(GoogleDriveReplyPayload))]
    internal sealed partial class GoogleDriveJsonSerializerContext : JsonSerializerContext {
    }

    internal static class GoogleDriveJson {
        internal static JsonNode? ToNode<T>(T value, JsonTypeInfo<T> typeInfo) {
            return JsonSerializer.SerializeToNode(value, typeInfo);
        }

        internal static string Serialize<T>(T value, JsonTypeInfo<T> typeInfo) {
            return JsonSerializer.Serialize(value, typeInfo);
        }
    }

    internal sealed class GoogleDriveFilePayload {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("mimeType")]
        public string? MimeType { get; set; }

        [JsonPropertyName("parents")]
        public string[]? Parents { get; set; }
    }

    internal sealed class GoogleDrivePermissionPayload {
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("role")]
        public string? Role { get; set; }

        [JsonPropertyName("emailAddress")]
        public string? EmailAddress { get; set; }

        [JsonPropertyName("domain")]
        public string? Domain { get; set; }

        [JsonPropertyName("allowFileDiscovery")]
        public bool? AllowFileDiscovery { get; set; }
    }

    internal sealed class GoogleDriveCommentPayload {
        [JsonPropertyName("content")]
        public string? Content { get; set; }

        [JsonPropertyName("anchor")]
        public string? Anchor { get; set; }
    }

    internal sealed class GoogleDriveReplyPayload {
        [JsonPropertyName("content")]
        public string? Content { get; set; }

        [JsonPropertyName("action")]
        public string? Action { get; set; }
    }
}
