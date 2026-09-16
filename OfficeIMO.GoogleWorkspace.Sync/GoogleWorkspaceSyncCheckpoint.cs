namespace OfficeIMO.GoogleWorkspace.Sync {
    /// <summary>Minimal durable state needed to resume Drive change tracking and relate local items to Google files.</summary>
    public sealed class GoogleWorkspaceSyncCheckpoint {
        /// <summary>Gets or sets the user change-feed token; a missing token is initialized at the current cursor.</summary>
        public string? UserChangeToken { get; set; }
        /// <summary>Gets the mutable map of shared-drive identifiers to their change-feed tokens.</summary>
        public IDictionary<string, string> SharedDriveChangeTokens { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
        /// <summary>Gets the mutable map of caller-owned item keys to stable local/Google identity evidence.</summary>
        public IDictionary<string, GoogleWorkspaceSyncIdentity> Identities { get; } = new Dictionary<string, GoogleWorkspaceSyncIdentity>(StringComparer.Ordinal);

        /// <summary>Copies the token maps and each identity so the copy can be advanced independently.</summary>
        public GoogleWorkspaceSyncCheckpoint Clone() {
            var clone = new GoogleWorkspaceSyncCheckpoint { UserChangeToken = UserChangeToken };
            foreach (KeyValuePair<string, string> pair in SharedDriveChangeTokens) clone.SharedDriveChangeTokens[pair.Key] = pair.Value;
            foreach (KeyValuePair<string, GoogleWorkspaceSyncIdentity> pair in Identities) clone.Identities[pair.Key] = pair.Value.Clone();
            return clone;
        }
    }

    /// <summary>Stable identity evidence only; content and document state remain owned by the caller.</summary>
    public sealed class GoogleWorkspaceSyncIdentity {
        /// <summary>Gets or sets the caller's local source identifier.</summary>
        public string? SourceId { get; set; }
        /// <summary>Gets or sets the corresponding Google Drive file identifier.</summary>
        public string? GoogleFileId { get; set; }
        /// <summary>Gets or sets the observed MIME type.</summary>
        public string? MimeType { get; set; }
        /// <summary>Gets or sets the observed revision identifier, when available.</summary>
        public string? RevisionId { get; set; }
        /// <summary>Gets or sets the observed Drive version, when available.</summary>
        public long? DriveVersion { get; set; }

        internal GoogleWorkspaceSyncIdentity Clone() => new GoogleWorkspaceSyncIdentity {
            SourceId = SourceId,
            GoogleFileId = GoogleFileId,
            MimeType = MimeType,
            RevisionId = RevisionId,
            DriveVersion = DriveVersion,
        };
    }
}
