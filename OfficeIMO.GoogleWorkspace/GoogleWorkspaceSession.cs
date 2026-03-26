namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Represents a configured Google Workspace session used by exporter packages.
    /// </summary>
    public sealed class GoogleWorkspaceSession {
        public GoogleWorkspaceSession(
            IGoogleWorkspaceCredentialSource credentialSource,
            GoogleWorkspaceSessionOptions? options = null) {
            CredentialSource = credentialSource ?? throw new ArgumentNullException(nameof(credentialSource));
            Options = options ?? new GoogleWorkspaceSessionOptions();
        }

        public IGoogleWorkspaceCredentialSource CredentialSource { get; }
        public GoogleWorkspaceSessionOptions Options { get; }

        public GoogleDriveFileLocation ResolveLocationDefaults(GoogleDriveFileLocation? location) {
            location ??= new GoogleDriveFileLocation();

            return new GoogleDriveFileLocation {
                DriveId = string.IsNullOrWhiteSpace(location.DriveId) ? Options.DefaultDriveId : location.DriveId,
                FolderId = string.IsNullOrWhiteSpace(location.FolderId) ? Options.DefaultFolderId : location.FolderId,
                ExistingFileId = location.ExistingFileId,
                SharedDriveAware = location.SharedDriveAware,
            };
        }

        public Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(
            IEnumerable<string> scopes,
            CancellationToken cancellationToken = default) {
            return CredentialSource.AcquireAccessTokenAsync(scopes, cancellationToken);
        }
    }
}
