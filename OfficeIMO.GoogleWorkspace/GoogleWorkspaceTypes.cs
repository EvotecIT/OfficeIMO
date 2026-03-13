namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Abstraction over the mechanism that acquires Google access tokens.
    /// </summary>
    public interface IGoogleWorkspaceCredentialSource {
        Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(
            IEnumerable<string> scopes,
            CancellationToken cancellationToken = default);
    }

    /// <summary>
    /// Represents an acquired Google OAuth access token.
    /// </summary>
    public sealed class GoogleWorkspaceAccessToken {
        public GoogleWorkspaceAccessToken(
            string accessToken,
            DateTimeOffset expiresAt,
            IReadOnlyList<string>? scopes = null) {
            AccessToken = accessToken ?? throw new ArgumentNullException(nameof(accessToken));
            ExpiresAt = expiresAt;
            Scopes = scopes ?? Array.Empty<string>();
        }

        public string AccessToken { get; }
        public DateTimeOffset ExpiresAt { get; }
        public IReadOnlyList<string> Scopes { get; }
        public bool IsExpired(DateTimeOffset now) => now >= ExpiresAt;
    }

    /// <summary>
    /// Describes the Drive target location for created or updated files.
    /// </summary>
    public sealed class GoogleDriveFileLocation {
        public string? DriveId { get; set; }
        public string? FolderId { get; set; }
        public string? ExistingFileId { get; set; }
        public bool SharedDriveAware { get; set; } = true;
    }

    /// <summary>
    /// Common Drive metadata returned by Google Workspace exporters.
    /// </summary>
    public class GoogleDriveFileReference {
        public string? FileId { get; set; }
        public string? Name { get; set; }
        public string? WebViewLink { get; set; }
        public string? MimeType { get; set; }
        public GoogleDriveFileLocation? Location { get; set; }
    }
}
