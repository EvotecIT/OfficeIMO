namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Abstraction over the mechanism that acquires Google access tokens.
    /// </summary>
    public interface IGoogleWorkspaceCredentialSource {
        /// <summary>Acquires an access token that is valid for every requested OAuth scope.</summary>
        /// <param name="scopes">OAuth scopes required by the pending operation.</param>
        /// <param name="cancellationToken">Token used to cancel acquisition.</param>
        /// <returns>A task that produces the acquired token and its scope and identity evidence.</returns>
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
        /// <summary>Creates verified credential evidence for an account and its granted scopes.</summary>
        /// <param name="account">Provider-verified account identity.</param>
        /// <param name="scopes">Provider-verified granted OAuth scopes.</param>
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

        /// <summary>Gets the normalized provider-verified account identity.</summary>
        public string Account { get; }
        /// <summary>Gets the nonempty, distinct granted scopes in ordinal order.</summary>
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

        /// <summary>Gets the OAuth bearer token.</summary>
        public string AccessToken { get; }
        /// <summary>Gets the instant after which the token must not be used.</summary>
        public DateTimeOffset ExpiresAt { get; }
        /// <summary>Gets the normalized OAuth scopes bound to the token by its source.</summary>
        public IReadOnlyList<string> Scopes { get; }
        /// <summary>Gets the credential-source account label, when supplied. Use <see cref="CredentialBinding"/> for verified evidence.</summary>
        public string? Account { get; }
        /// <summary>Gets provider-verified account and scope evidence, when supplied by the credential source.</summary>
        public GoogleWorkspaceCredentialBinding? CredentialBinding { get; }
        /// <summary>Determines whether the token has expired at the supplied instant.</summary>
        /// <param name="now">Instant to compare with <see cref="ExpiresAt"/>.</param>
        /// <returns><see langword="true"/> when <paramref name="now"/> is at or after the expiry instant.</returns>
        public bool IsExpired(DateTimeOffset now) => now >= ExpiresAt;
    }

    /// <summary>
    /// Describes the Drive target location for created or updated files.
    /// </summary>
    public sealed class GoogleDriveFileLocation {
        /// <summary>Gets or sets the shared-drive identifier, or <see langword="null"/> for My Drive.</summary>
        public string? DriveId { get; set; }
        /// <summary>Gets or sets the parent folder identifier for newly created files.</summary>
        public string? FolderId { get; set; }
        /// <summary>Gets or sets the identifier of an existing file to update instead of creating one.</summary>
        public string? ExistingFileId { get; set; }
        /// <summary>Gets or sets whether Drive requests include shared-drive support flags. The default is <see langword="true"/>.</summary>
        public bool SharedDriveAware { get; set; } = true;
    }

    /// <summary>
    /// Common Drive metadata returned by Google Workspace exporters.
    /// </summary>
    public class GoogleDriveFileReference {
        /// <summary>Gets or sets the Google Drive file identifier.</summary>
        public string? FileId { get; set; }
        /// <summary>Gets or sets the display name returned by Google Drive.</summary>
        public string? Name { get; set; }
        /// <summary>Gets or sets the browser URL for viewing the file.</summary>
        public string? WebViewLink { get; set; }
        /// <summary>Gets or sets the file's Google Drive MIME type.</summary>
        public string? MimeType { get; set; }
        /// <summary>Gets or sets the resolved Drive location used for the operation.</summary>
        public GoogleDriveFileLocation? Location { get; set; }
    }
}
