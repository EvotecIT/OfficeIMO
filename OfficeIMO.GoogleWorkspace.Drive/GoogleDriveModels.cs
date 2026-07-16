using System.Text.Json.Serialization;

namespace OfficeIMO.GoogleWorkspace.Drive {
    public static class GoogleDriveMimeTypes {
        public const string Folder = "application/vnd.google-apps.folder";
        public const string Document = "application/vnd.google-apps.document";
        public const string Spreadsheet = "application/vnd.google-apps.spreadsheet";
        public const string Presentation = "application/vnd.google-apps.presentation";
        public const string MicrosoftWord = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        public const string MicrosoftExcel = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        public const string MicrosoftPowerPoint = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    }

    public sealed class GoogleDriveFile {
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("mimeType")]
        public string? MimeType { get; set; }

        [JsonPropertyName("driveId")]
        public string? DriveId { get; set; }

        [JsonPropertyName("parents")]
        public List<string> Parents { get; set; } = new List<string>();

        [JsonPropertyName("webViewLink")]
        public string? WebViewLink { get; set; }

        [JsonPropertyName("webContentLink")]
        public string? WebContentLink { get; set; }

        [JsonPropertyName("modifiedTime")]
        public DateTimeOffset? ModifiedTime { get; set; }

        [JsonPropertyName("createdTime")]
        public DateTimeOffset? CreatedTime { get; set; }

        [JsonPropertyName("version")]
        [JsonNumberHandling(JsonNumberHandling.AllowReadingFromString)]
        public long? Version { get; set; }

        [JsonPropertyName("size")]
        [JsonNumberHandling(JsonNumberHandling.AllowReadingFromString)]
        public long? Size { get; set; }

        [JsonPropertyName("trashed")]
        public bool Trashed { get; set; }

        [JsonPropertyName("capabilities")]
        public GoogleDriveFileCapabilities? Capabilities { get; set; }
    }

    public sealed class GoogleDriveFileCapabilities {
        [JsonPropertyName("canDownload")]
        public bool CanDownload { get; set; }

        [JsonPropertyName("canEdit")]
        public bool CanEdit { get; set; }

        [JsonPropertyName("canMoveItemWithinDrive")]
        public bool CanMoveItemWithinDrive { get; set; }

        [JsonPropertyName("canMoveItemOutOfDrive")]
        public bool CanMoveItemOutOfDrive { get; set; }

        [JsonPropertyName("canDelete")]
        public bool CanDelete { get; set; }

        [JsonPropertyName("canShare")]
        public bool CanShare { get; set; }

        [JsonPropertyName("canComment")]
        public bool CanComment { get; set; }
    }

    public sealed class GoogleDriveFileList {
        [JsonPropertyName("files")]
        public List<GoogleDriveFile> Files { get; set; } = new List<GoogleDriveFile>();

        [JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }

        [JsonPropertyName("incompleteSearch")]
        public bool IncompleteSearch { get; set; }
    }

    public sealed class GoogleSharedDrive {
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("hidden")]
        public bool Hidden { get; set; }

        [JsonPropertyName("createdTime")]
        public DateTimeOffset? CreatedTime { get; set; }

        [JsonPropertyName("capabilities")]
        public GoogleSharedDriveCapabilities? Capabilities { get; set; }
    }

    public sealed class GoogleSharedDriveCapabilities {
        [JsonPropertyName("canAddChildren")]
        public bool CanAddChildren { get; set; }

        [JsonPropertyName("canManageMembers")]
        public bool CanManageMembers { get; set; }

        [JsonPropertyName("canRename")]
        public bool CanRename { get; set; }
    }

    public sealed class GoogleDriveAboutFormats {
        [JsonPropertyName("importFormats")]
        public Dictionary<string, List<string>> ImportFormats { get; set; } = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        [JsonPropertyName("exportFormats")]
        public Dictionary<string, List<string>> ExportFormats { get; set; } = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class GoogleDrivePermission {
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("role")]
        public string? Role { get; set; }

        [JsonPropertyName("emailAddress")]
        public string? EmailAddress { get; set; }

        [JsonPropertyName("domain")]
        public string? Domain { get; set; }

        [JsonPropertyName("displayName")]
        public string? DisplayName { get; set; }

        [JsonPropertyName("allowFileDiscovery")]
        public bool? AllowFileDiscovery { get; set; }
    }

    public sealed class GoogleDrivePermissionList {
        [JsonPropertyName("permissions")]
        public List<GoogleDrivePermission> Permissions { get; set; } = new List<GoogleDrivePermission>();

        [JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }
    }

    public sealed class GoogleDriveComment {
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("content")]
        public string? Content { get; set; }

        [JsonPropertyName("anchor")]
        public string? Anchor { get; set; }

        [JsonPropertyName("resolved")]
        public bool Resolved { get; set; }

        [JsonPropertyName("deleted")]
        public bool Deleted { get; set; }

        [JsonPropertyName("createdTime")]
        public DateTimeOffset? CreatedTime { get; set; }

        [JsonPropertyName("modifiedTime")]
        public DateTimeOffset? ModifiedTime { get; set; }

        [JsonPropertyName("replies")]
        public List<GoogleDriveReply> Replies { get; set; } = new List<GoogleDriveReply>();
    }

    public sealed class GoogleDriveCommentList {
        [JsonPropertyName("comments")]
        public List<GoogleDriveComment> Comments { get; set; } = new List<GoogleDriveComment>();

        [JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }
    }

    public sealed class GoogleDriveReply {
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("content")]
        public string? Content { get; set; }

        [JsonPropertyName("action")]
        public string? Action { get; set; }

        [JsonPropertyName("deleted")]
        public bool Deleted { get; set; }

        [JsonPropertyName("createdTime")]
        public DateTimeOffset? CreatedTime { get; set; }
    }

    public sealed class GoogleDriveRevision {
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("modifiedTime")]
        public DateTimeOffset? ModifiedTime { get; set; }

        [JsonPropertyName("keepForever")]
        public bool KeepForever { get; set; }

        [JsonPropertyName("published")]
        public bool Published { get; set; }

        [JsonPropertyName("size")]
        [JsonNumberHandling(JsonNumberHandling.AllowReadingFromString)]
        public long? Size { get; set; }
    }

    public sealed class GoogleDriveRevisionList {
        [JsonPropertyName("revisions")]
        public List<GoogleDriveRevision> Revisions { get; set; } = new List<GoogleDriveRevision>();

        [JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }
    }

    public sealed class GoogleDriveChange {
        [JsonPropertyName("fileId")]
        public string? FileId { get; set; }

        [JsonPropertyName("removed")]
        public bool Removed { get; set; }

        [JsonPropertyName("changeType")]
        public string? ChangeType { get; set; }

        [JsonPropertyName("file")]
        public GoogleDriveFile? File { get; set; }

        [JsonPropertyName("driveId")]
        public string? DriveId { get; set; }
    }

    public sealed class GoogleDriveChangeList {
        [JsonPropertyName("changes")]
        public List<GoogleDriveChange> Changes { get; set; } = new List<GoogleDriveChange>();

        [JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }

        [JsonPropertyName("newStartPageToken")]
        public string? NewStartPageToken { get; set; }
    }

    public sealed class GoogleDriveStartPageToken {
        [JsonPropertyName("startPageToken")]
        public string? Value { get; set; }
    }
}
