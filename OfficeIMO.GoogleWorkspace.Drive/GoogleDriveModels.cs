using System.Text.Json.Serialization;

namespace OfficeIMO.GoogleWorkspace.Drive {
    /// <summary>Google Drive and Microsoft Office MIME types used by Drive operations.</summary>
    public static class GoogleDriveMimeTypes {
        /// <summary>Google Drive folder MIME type.</summary>
        public const string Folder = "application/vnd.google-apps.folder";
        /// <summary>Google Docs document MIME type.</summary>
        public const string Document = "application/vnd.google-apps.document";
        /// <summary>Google Sheets spreadsheet MIME type.</summary>
        public const string Spreadsheet = "application/vnd.google-apps.spreadsheet";
        /// <summary>Google Slides presentation MIME type.</summary>
        public const string Presentation = "application/vnd.google-apps.presentation";
        /// <summary>Office Open XML Word document MIME type.</summary>
        public const string MicrosoftWord = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        /// <summary>Office Open XML Excel workbook MIME type.</summary>
        public const string MicrosoftExcel = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        /// <summary>Office Open XML PowerPoint presentation MIME type.</summary>
        public const string MicrosoftPowerPoint = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    }

    /// <summary>Metadata and capabilities returned for a Google Drive file.</summary>
    public sealed class GoogleDriveFile {
        /// <summary>Gets or sets the Drive file identifier.</summary>
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        /// <summary>Gets or sets the display name.</summary>
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        /// <summary>Gets or sets the Drive MIME type.</summary>
        [JsonPropertyName("mimeType")]
        public string? MimeType { get; set; }

        /// <summary>Gets or sets the containing shared-drive identifier.</summary>
        [JsonPropertyName("driveId")]
        public string? DriveId { get; set; }

        /// <summary>Gets or sets parent-folder identifiers.</summary>
        [JsonPropertyName("parents")]
        public List<string> Parents { get; set; } = new List<string>();

        /// <summary>Gets or sets the browser URL for viewing the file.</summary>
        [JsonPropertyName("webViewLink")]
        public string? WebViewLink { get; set; }

        /// <summary>Gets or sets the browser URL for downloading binary content.</summary>
        [JsonPropertyName("webContentLink")]
        public string? WebContentLink { get; set; }

        /// <summary>Gets or sets the most recent modification time.</summary>
        [JsonPropertyName("modifiedTime")]
        public DateTimeOffset? ModifiedTime { get; set; }

        /// <summary>Gets or sets the creation time.</summary>
        [JsonPropertyName("createdTime")]
        public DateTimeOffset? CreatedTime { get; set; }

        /// <summary>Gets or sets the monotonically increasing Drive file version.</summary>
        [JsonPropertyName("version")]
        [JsonNumberHandling(JsonNumberHandling.AllowReadingFromString)]
        public long? Version { get; set; }

        /// <summary>Gets or sets the binary content size in bytes, when applicable.</summary>
        [JsonPropertyName("size")]
        [JsonNumberHandling(JsonNumberHandling.AllowReadingFromString)]
        public long? Size { get; set; }

        /// <summary>Gets or sets whether the file is in the trash.</summary>
        [JsonPropertyName("trashed")]
        public bool Trashed { get; set; }

        /// <summary>Gets or sets actions permitted for the authenticated caller.</summary>
        [JsonPropertyName("capabilities")]
        public GoogleDriveFileCapabilities? Capabilities { get; set; }
    }

    /// <summary>Actions the authenticated caller may perform on a Drive file.</summary>
    public sealed class GoogleDriveFileCapabilities {
        /// <summary>Gets or sets whether file content may be downloaded.</summary>
        [JsonPropertyName("canDownload")]
        public bool CanDownload { get; set; }

        /// <summary>Gets or sets whether file content or metadata may be edited.</summary>
        [JsonPropertyName("canEdit")]
        public bool CanEdit { get; set; }

        /// <summary>Gets or sets whether the item may be moved within its current drive.</summary>
        [JsonPropertyName("canMoveItemWithinDrive")]
        public bool CanMoveItemWithinDrive { get; set; }

        /// <summary>Gets or sets whether the item may be moved out of its current drive.</summary>
        [JsonPropertyName("canMoveItemOutOfDrive")]
        public bool CanMoveItemOutOfDrive { get; set; }

        /// <summary>Gets or sets whether the file may be deleted.</summary>
        [JsonPropertyName("canDelete")]
        public bool CanDelete { get; set; }

        /// <summary>Gets or sets whether permissions may be changed.</summary>
        [JsonPropertyName("canShare")]
        public bool CanShare { get; set; }

        /// <summary>Gets or sets whether comments may be added.</summary>
        [JsonPropertyName("canComment")]
        public bool CanComment { get; set; }
    }

    /// <summary>One page of Google Drive files.</summary>
    public sealed class GoogleDriveFileList {
        /// <summary>Gets or sets files returned on this page.</summary>
        [JsonPropertyName("files")]
        public List<GoogleDriveFile> Files { get; set; } = new List<GoogleDriveFile>();

        /// <summary>Gets or sets the token used to request the next page.</summary>
        [JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }

        /// <summary>Gets or sets whether Drive could not search every requested corpus.</summary>
        [JsonPropertyName("incompleteSearch")]
        public bool IncompleteSearch { get; set; }
    }

    /// <summary>Metadata for a Google shared drive.</summary>
    public sealed class GoogleSharedDrive {
        /// <summary>Gets or sets the shared-drive identifier.</summary>
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        /// <summary>Gets or sets the shared-drive display name.</summary>
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        /// <summary>Gets or sets whether the shared drive is hidden from the default view.</summary>
        [JsonPropertyName("hidden")]
        public bool Hidden { get; set; }

        /// <summary>Gets or sets when the shared drive was created.</summary>
        [JsonPropertyName("createdTime")]
        public DateTimeOffset? CreatedTime { get; set; }

        /// <summary>Gets or sets actions permitted for the authenticated caller.</summary>
        [JsonPropertyName("capabilities")]
        public GoogleSharedDriveCapabilities? Capabilities { get; set; }
    }

    /// <summary>Actions the authenticated caller may perform on a shared drive.</summary>
    public sealed class GoogleSharedDriveCapabilities {
        /// <summary>Gets or sets whether items may be added to the drive.</summary>
        [JsonPropertyName("canAddChildren")]
        public bool CanAddChildren { get; set; }

        /// <summary>Gets or sets whether shared-drive membership may be managed.</summary>
        [JsonPropertyName("canManageMembers")]
        public bool CanManageMembers { get; set; }

        /// <summary>Gets or sets whether the shared drive may be renamed.</summary>
        [JsonPropertyName("canRename")]
        public bool CanRename { get; set; }
    }

    /// <summary>Import and export format mappings reported by the Drive About resource.</summary>
    public sealed class GoogleDriveAboutFormats {
        /// <summary>Gets or sets source MIME types and the Google formats into which Drive can import them.</summary>
        [JsonPropertyName("importFormats")]
        public Dictionary<string, List<string>> ImportFormats { get; set; } = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Gets or sets Google MIME types and the external formats to which Drive can export them.</summary>
        [JsonPropertyName("exportFormats")]
        public Dictionary<string, List<string>> ExportFormats { get; set; } = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>A permission granted on a Drive file or shared drive.</summary>
    public sealed class GoogleDrivePermission {
        /// <summary>Gets or sets the permission identifier.</summary>
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        /// <summary>Gets or sets the grantee type, such as user, group, domain, or anyone.</summary>
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        /// <summary>Gets or sets the granted role, such as reader, commenter, writer, or owner.</summary>
        [JsonPropertyName("role")]
        public string? Role { get; set; }

        /// <summary>Gets or sets the user or group email address.</summary>
        [JsonPropertyName("emailAddress")]
        public string? EmailAddress { get; set; }

        /// <summary>Gets or sets the grantee domain for domain permissions.</summary>
        [JsonPropertyName("domain")]
        public string? Domain { get; set; }

        /// <summary>Gets or sets the grantee's display name.</summary>
        [JsonPropertyName("displayName")]
        public string? DisplayName { get; set; }

        /// <summary>Gets or sets whether the permission makes the file discoverable through search.</summary>
        [JsonPropertyName("allowFileDiscovery")]
        public bool? AllowFileDiscovery { get; set; }
    }

    /// <summary>One page of permissions for a Drive resource.</summary>
    public sealed class GoogleDrivePermissionList {
        /// <summary>Gets or sets permissions returned on this page.</summary>
        [JsonPropertyName("permissions")]
        public List<GoogleDrivePermission> Permissions { get; set; } = new List<GoogleDrivePermission>();

        /// <summary>Gets or sets the token used to request the next page.</summary>
        [JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }
    }

    /// <summary>A comment attached to a Drive file.</summary>
    public sealed class GoogleDriveComment {
        /// <summary>Gets or sets the comment identifier.</summary>
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        /// <summary>Gets or sets the comment's plain-text content.</summary>
        [JsonPropertyName("content")]
        public string? Content { get; set; }

        /// <summary>Gets or sets the JSON anchor locating the commented region.</summary>
        [JsonPropertyName("anchor")]
        public string? Anchor { get; set; }

        /// <summary>Gets or sets whether the comment thread is resolved.</summary>
        [JsonPropertyName("resolved")]
        public bool Resolved { get; set; }

        /// <summary>Gets or sets whether the comment has been deleted.</summary>
        [JsonPropertyName("deleted")]
        public bool Deleted { get; set; }

        /// <summary>Gets or sets when the comment was created.</summary>
        [JsonPropertyName("createdTime")]
        public DateTimeOffset? CreatedTime { get; set; }

        /// <summary>Gets or sets when the comment was most recently modified.</summary>
        [JsonPropertyName("modifiedTime")]
        public DateTimeOffset? ModifiedTime { get; set; }

        /// <summary>Gets or sets replies in the comment thread.</summary>
        [JsonPropertyName("replies")]
        public List<GoogleDriveReply> Replies { get; set; } = new List<GoogleDriveReply>();
    }

    /// <summary>One page of comments for a Drive file.</summary>
    public sealed class GoogleDriveCommentList {
        /// <summary>Gets or sets comments returned on this page.</summary>
        [JsonPropertyName("comments")]
        public List<GoogleDriveComment> Comments { get; set; } = new List<GoogleDriveComment>();

        /// <summary>Gets or sets the token used to request the next page.</summary>
        [JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }
    }

    /// <summary>A reply in a Drive comment thread.</summary>
    public sealed class GoogleDriveReply {
        /// <summary>Gets or sets the reply identifier.</summary>
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        /// <summary>Gets or sets the reply's plain-text content.</summary>
        [JsonPropertyName("content")]
        public string? Content { get; set; }

        /// <summary>Gets or sets the action applied to the parent comment, such as resolve or reopen.</summary>
        [JsonPropertyName("action")]
        public string? Action { get; set; }

        /// <summary>Gets or sets whether the reply has been deleted.</summary>
        [JsonPropertyName("deleted")]
        public bool Deleted { get; set; }

        /// <summary>Gets or sets when the reply was created.</summary>
        [JsonPropertyName("createdTime")]
        public DateTimeOffset? CreatedTime { get; set; }
    }

    /// <summary>Metadata for one historical Drive file revision.</summary>
    public sealed class GoogleDriveRevision {
        /// <summary>Gets or sets the revision identifier.</summary>
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        /// <summary>Gets or sets when the revision was created or modified.</summary>
        [JsonPropertyName("modifiedTime")]
        public DateTimeOffset? ModifiedTime { get; set; }

        /// <summary>Gets or sets whether the revision is retained indefinitely.</summary>
        [JsonPropertyName("keepForever")]
        public bool KeepForever { get; set; }

        /// <summary>Gets or sets whether the revision is published.</summary>
        [JsonPropertyName("published")]
        public bool Published { get; set; }

        /// <summary>Gets or sets the revision size in bytes, when available.</summary>
        [JsonPropertyName("size")]
        [JsonNumberHandling(JsonNumberHandling.AllowReadingFromString)]
        public long? Size { get; set; }
    }

    /// <summary>One page of Drive file revisions.</summary>
    public sealed class GoogleDriveRevisionList {
        /// <summary>Gets or sets revisions returned on this page.</summary>
        [JsonPropertyName("revisions")]
        public List<GoogleDriveRevision> Revisions { get; set; } = new List<GoogleDriveRevision>();

        /// <summary>Gets or sets the token used to request the next page.</summary>
        [JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }
    }

    /// <summary>One change recorded by the Google Drive change feed.</summary>
    public sealed class GoogleDriveChange {
        /// <summary>Gets or sets the affected file identifier.</summary>
        [JsonPropertyName("fileId")]
        public string? FileId { get; set; }

        /// <summary>Gets or sets whether the resource was removed from the requested change corpus.</summary>
        [JsonPropertyName("removed")]
        public bool Removed { get; set; }

        /// <summary>Gets or sets whether the changed resource is a file or shared drive.</summary>
        [JsonPropertyName("changeType")]
        public string? ChangeType { get; set; }

        /// <summary>Gets or sets current file metadata when the file remains accessible.</summary>
        [JsonPropertyName("file")]
        public GoogleDriveFile? File { get; set; }

        /// <summary>Gets or sets the affected shared-drive identifier.</summary>
        [JsonPropertyName("driveId")]
        public string? DriveId { get; set; }
    }

    /// <summary>One page of Drive changes and continuation state.</summary>
    public sealed class GoogleDriveChangeList {
        /// <summary>Gets or sets changes returned on this page.</summary>
        [JsonPropertyName("changes")]
        public List<GoogleDriveChange> Changes { get; set; } = new List<GoogleDriveChange>();

        /// <summary>Gets or sets the token used to request the next page.</summary>
        [JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }

        /// <summary>Gets or sets the new starting token after the current change stream is exhausted.</summary>
        [JsonPropertyName("newStartPageToken")]
        public string? NewStartPageToken { get; set; }
    }

    /// <summary>Starting token for subsequent Google Drive change-feed reads.</summary>
    public sealed class GoogleDriveStartPageToken {
        /// <summary>Gets or sets the opaque change-feed token.</summary>
        [JsonPropertyName("startPageToken")]
        public string? Value { get; set; }
    }
}
