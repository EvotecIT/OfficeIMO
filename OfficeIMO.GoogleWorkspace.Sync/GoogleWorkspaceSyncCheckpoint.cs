namespace OfficeIMO.GoogleWorkspace.Sync {
    /// <summary>Minimal durable state needed to resume Drive change tracking and relate local items to Google files.</summary>
    public sealed class GoogleWorkspaceSyncCheckpoint {
        public string? UserChangeToken { get; set; }
        public IDictionary<string, string> SharedDriveChangeTokens { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
        public IDictionary<string, GoogleWorkspaceSyncIdentity> Identities { get; } = new Dictionary<string, GoogleWorkspaceSyncIdentity>(StringComparer.Ordinal);

        public GoogleWorkspaceSyncCheckpoint Clone() {
            var clone = new GoogleWorkspaceSyncCheckpoint { UserChangeToken = UserChangeToken };
            foreach (KeyValuePair<string, string> pair in SharedDriveChangeTokens) clone.SharedDriveChangeTokens[pair.Key] = pair.Value;
            foreach (KeyValuePair<string, GoogleWorkspaceSyncIdentity> pair in Identities) clone.Identities[pair.Key] = pair.Value.Clone();
            return clone;
        }
    }

    /// <summary>Stable identity evidence only; content and document state remain owned by the caller.</summary>
    public sealed class GoogleWorkspaceSyncIdentity {
        public string? SourceId { get; set; }
        public string? GoogleFileId { get; set; }
        public string? MimeType { get; set; }
        public string? RevisionId { get; set; }
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
