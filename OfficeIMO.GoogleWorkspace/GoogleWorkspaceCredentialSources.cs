namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Credential source that always returns the same already-acquired access token.
    /// </summary>
    public sealed class StaticAccessTokenCredentialSource : IGoogleWorkspaceCredentialSource {
        private readonly string _accessToken;
        private readonly DateTimeOffset _expiresAt;
        private readonly IReadOnlyList<string>? _scopes;
        private readonly string? _account;

        /// <summary>Creates an unverified static token source. This exact overload is retained for binary compatibility.</summary>
        public StaticAccessTokenCredentialSource(
            string accessToken,
            DateTimeOffset? expiresAt = null,
            IReadOnlyList<string>? scopes = null)
            : this(accessToken, expiresAt, scopes, null) { }

        /// <summary>Creates an unverified static token source with an informational account label.</summary>
        public StaticAccessTokenCredentialSource(
            string accessToken,
            DateTimeOffset? expiresAt,
            IReadOnlyList<string>? scopes,
            string? account) {
            _accessToken = accessToken ?? throw new ArgumentNullException(nameof(accessToken));
            _expiresAt = expiresAt ?? DateTimeOffset.UtcNow.AddMinutes(30);
            _scopes = scopes;
            _account = account;
        }

        public Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(
            IEnumerable<string> scopes,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<string> effectiveScopes = _scopes ?? scopes?.ToArray() ?? Array.Empty<string>();

            return Task.FromResult(new GoogleWorkspaceAccessToken(
                _accessToken, _expiresAt, effectiveScopes, _account));
        }
    }

    /// <summary>
    /// Credential source that delegates token acquisition to application-provided code.
    /// </summary>
    public sealed class DelegateGoogleWorkspaceCredentialSource : IGoogleWorkspaceCredentialSource {
        private readonly Func<IReadOnlyList<string>, CancellationToken, Task<GoogleWorkspaceAccessToken>> _acquireAccessToken;

        public DelegateGoogleWorkspaceCredentialSource(
            Func<IReadOnlyList<string>, CancellationToken, Task<GoogleWorkspaceAccessToken>> acquireAccessToken) {
            _acquireAccessToken = acquireAccessToken ?? throw new ArgumentNullException(nameof(acquireAccessToken));
        }

        public Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(
            IEnumerable<string> scopes,
            CancellationToken cancellationToken = default) {
            IReadOnlyList<string> requestedScopes = scopes?.ToArray() ?? Array.Empty<string>();
            return _acquireAccessToken(requestedScopes, cancellationToken);
        }
    }
}
