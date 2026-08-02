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
    /// Account and grant evidence obtained from a trusted credential or token-verification source.
    /// </summary>
    /// <remarks>
    /// Construct this only from provider-issued evidence. Caller-entered account names and requested scopes
    /// are policy inputs, not credential evidence.
    /// </remarks>
    public sealed class GoogleWorkspaceCredentialBinding {
        public GoogleWorkspaceCredentialBinding(string account, IReadOnlyList<string> scopes) {
            if (string.IsNullOrWhiteSpace(account)) throw new ArgumentException("A verified account identity is required.", nameof(account));
            Account = account.Trim();
            Scopes = Array.AsReadOnly((scopes ?? throw new ArgumentNullException(nameof(scopes)))
                .Where(scope => !string.IsNullOrWhiteSpace(scope))
                .Select(scope => scope!.Trim())
                .Distinct(StringComparer.Ordinal)
                .OrderBy(scope => scope, StringComparer.Ordinal)
                .ToArray());
            if (Scopes.Count == 0) throw new ArgumentException("At least one verified scope is required.", nameof(scopes));
        }

        public string Account { get; }
        public IReadOnlyList<string> Scopes { get; }
    }

    /// <summary>Resolves provider-verified account and grant evidence for an acquired access token.</summary>
    public delegate Task<GoogleWorkspaceCredentialBinding> GoogleWorkspaceCredentialBindingResolver(
        string accessToken,
        CancellationToken cancellationToken);

    /// <summary>
    /// Represents an acquired Google OAuth access token.
    /// </summary>
    public sealed class GoogleWorkspaceAccessToken {
        /// <summary>Creates a token without verified account or grant evidence.</summary>
        public GoogleWorkspaceAccessToken(
            string accessToken,
            DateTimeOffset expiresAt,
            IReadOnlyList<string>? scopes = null)
            : this(accessToken, expiresAt, scopes, null, null) { }

        /// <summary>
        /// Creates a token with an informational account label. The label is not credential evidence and
        /// does not make the token eligible for guarded mutations.
        /// </summary>
        public GoogleWorkspaceAccessToken(
            string accessToken,
            DateTimeOffset expiresAt,
            IReadOnlyList<string>? scopes,
            string? account)
            : this(accessToken, expiresAt, scopes, account, null) { }

        private GoogleWorkspaceAccessToken(
            string accessToken,
            DateTimeOffset expiresAt,
            IReadOnlyList<string>? scopes,
            string? account,
            GoogleWorkspaceCredentialBinding? credentialBinding) {
            if (string.IsNullOrWhiteSpace(accessToken)) throw new ArgumentException("Access token is required.", nameof(accessToken));
            AccessToken = accessToken;
            ExpiresAt = expiresAt;
            Scopes = Array.AsReadOnly((scopes ?? Array.Empty<string>())
                .Where(scope => !string.IsNullOrWhiteSpace(scope))
                .Select(scope => scope!.Trim())
                .Distinct(StringComparer.Ordinal)
                .ToArray());
            Account = string.IsNullOrWhiteSpace(account) ? null : account!.Trim();
            CredentialBinding = credentialBinding;
        }

        /// <summary>Creates a token bound to provider-verified account and scope evidence.</summary>
        public static GoogleWorkspaceAccessToken FromVerifiedCredential(
            string accessToken,
            DateTimeOffset expiresAt,
            GoogleWorkspaceCredentialBinding credentialBinding) {
            if (credentialBinding == null) throw new ArgumentNullException(nameof(credentialBinding));
            return new GoogleWorkspaceAccessToken(accessToken, expiresAt, credentialBinding.Scopes,
                credentialBinding.Account, credentialBinding);
        }

        public string AccessToken { get; }
        public DateTimeOffset ExpiresAt { get; }
        public IReadOnlyList<string> Scopes { get; }
        /// <summary>Gets the credential-source account label, when supplied. Use <see cref="CredentialBinding"/> for verified evidence.</summary>
        public string? Account { get; }
        /// <summary>Gets provider-verified account and scope evidence, when supplied by the credential source.</summary>
        public GoogleWorkspaceCredentialBinding? CredentialBinding { get; }
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
