namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Credential source that always returns the same already-acquired access token.
    /// </summary>
    public sealed class StaticAccessTokenCredentialSource : IGoogleWorkspaceCredentialSource {
        private readonly string _accessToken;
        private readonly DateTimeOffset _expiresAt;
        private readonly IReadOnlyList<string>? _scopes;

        public StaticAccessTokenCredentialSource(
            string accessToken,
            DateTimeOffset? expiresAt = null,
            IReadOnlyList<string>? scopes = null) {
            _accessToken = accessToken ?? throw new ArgumentNullException(nameof(accessToken));
            _expiresAt = expiresAt ?? DateTimeOffset.UtcNow.AddMinutes(30);
            _scopes = scopes;
        }

        public Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(
            IEnumerable<string> scopes,
            CancellationToken cancellationToken = default) {
            IReadOnlyList<string> effectiveScopes = _scopes ?? scopes?.ToArray() ?? Array.Empty<string>();

            return Task.FromResult(new GoogleWorkspaceAccessToken(
                _accessToken,
                _expiresAt,
                effectiveScopes));
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
