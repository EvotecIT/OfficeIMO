namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Represents a configured Google Workspace session used by exporter packages.
    /// </summary>
    public sealed class GoogleWorkspaceSession {
        private readonly System.Collections.Concurrent.ConcurrentDictionary<string, GoogleWorkspaceAccessToken>
            _verifiedTokens = new System.Collections.Concurrent.ConcurrentDictionary<string, GoogleWorkspaceAccessToken>(StringComparer.Ordinal);
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

        public async Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(
            IEnumerable<string> scopes,
            CancellationToken cancellationToken = default) {
            string[] requestedScopes = (scopes ?? Array.Empty<string>())
                .Where(scope => !string.IsNullOrWhiteSpace(scope))
                .Select(scope => scope!.Trim())
                .Distinct(StringComparer.Ordinal)
                .OrderBy(scope => scope, StringComparer.Ordinal)
                .ToArray();
            GoogleWorkspaceAccessToken token = await CredentialSource
                .AcquireAccessTokenAsync(requestedScopes, cancellationToken).ConfigureAwait(false);
            if (token == null) throw new InvalidOperationException("The Google credential source returned no access token.");
            if (token.IsExpired(DateTimeOffset.UtcNow)) throw new InvalidOperationException("The Google credential source returned an expired access token.");
            if (!new HashSet<string>(token.Scopes, StringComparer.Ordinal).IsSupersetOf(requestedScopes)) {
                throw new InvalidOperationException("The Google credential source did not bind the access token to every requested scope.");
            }
            if (!string.IsNullOrWhiteSpace(Options.ExpectedAccount)) {
                if (token.CredentialBinding == null) {
                    throw new InvalidOperationException("The Google credential source did not provide provider-verified account and scope evidence.");
                }
                if (!StringComparer.OrdinalIgnoreCase.Equals(token.CredentialBinding.Account, Options.ExpectedAccount)) {
                    throw new InvalidOperationException("The acquired Google credential account does not match the configured expected account.");
                }
                if (!new HashSet<string>(token.CredentialBinding.Scopes, StringComparer.Ordinal).IsSupersetOf(requestedScopes)) {
                    throw new InvalidOperationException("The provider-verified Google credential grants do not contain every requested scope.");
                }
            }
            _verifiedTokens[CreateBindingKey(token.AccessToken, requestedScopes)] = token;
            return token;
        }

        internal string VerifyMutationCredential(string accessToken, IReadOnlyCollection<string> requiredScopes) {
            if (!_verifiedTokens.TryGetValue(CreateBindingKey(accessToken, requiredScopes),
                    out GoogleWorkspaceAccessToken? token)
                || token.IsExpired(DateTimeOffset.UtcNow)) {
                throw new InvalidOperationException("Google mutations require an access token acquired and verified by this session.");
            }
            if (token.CredentialBinding == null) {
                throw new InvalidOperationException("Google mutations require provider-verified account and scope evidence.");
            }
            if (!new HashSet<string>(token.Scopes, StringComparer.Ordinal).IsSupersetOf(requiredScopes)) {
                throw new InvalidOperationException("The Google mutation scopes are not contained in the scopes bound to the acquired access token.");
            }
            if (!new HashSet<string>(token.CredentialBinding.Scopes, StringComparer.Ordinal).IsSupersetOf(requiredScopes)) {
                throw new InvalidOperationException("The Google mutation scopes are not contained in the provider-verified credential grants.");
            }
            return token.CredentialBinding.Account;
        }

        private static string CreateBindingKey(string accessToken, IEnumerable<string> scopes) =>
            accessToken + "\0" + string.Join("\n", scopes.OrderBy(scope => scope, StringComparer.Ordinal));
    }
}
