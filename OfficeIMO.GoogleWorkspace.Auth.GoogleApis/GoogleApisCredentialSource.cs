using Google.Apis.Auth.OAuth2;

namespace OfficeIMO.GoogleWorkspace.Auth.GoogleApis {
    /// <summary>
    /// Adapts credentials from Google.Apis.Auth to the dependency-light OfficeIMO credential contract.
    /// </summary>
    public sealed class GoogleApisCredentialSource : IGoogleWorkspaceCredentialSource {
        private static readonly TimeSpan DefaultLifetime = TimeSpan.FromMinutes(50);
        private readonly Func<IReadOnlyList<string>, ITokenAccess> _credentialFactory;
        private readonly TimeSpan _fallbackTokenLifetime;

        /// <summary>
        /// Creates an adapter around a Google credential, applying requested scopes when the credential requires scoping.
        /// </summary>
        public GoogleApisCredentialSource(
            GoogleCredential credential,
            TimeSpan? fallbackTokenLifetime = null) {
            if (credential == null) {
                throw new ArgumentNullException(nameof(credential));
            }

            _credentialFactory = scopes => credential.IsCreateScopedRequired
                ? credential.CreateScoped(scopes)
                : credential;
            _fallbackTokenLifetime = ValidateLifetime(fallbackTokenLifetime);
        }

        /// <summary>
        /// Creates an adapter around any Google token source, including <see cref="UserCredential"/>.
        /// </summary>
        public GoogleApisCredentialSource(
            ITokenAccess credential,
            TimeSpan? fallbackTokenLifetime = null) {
            if (credential == null) {
                throw new ArgumentNullException(nameof(credential));
            }

            _credentialFactory = _ => credential;
            _fallbackTokenLifetime = ValidateLifetime(fallbackTokenLifetime);
        }

        /// <inheritdoc />
        public async Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(
            IEnumerable<string> scopes,
            CancellationToken cancellationToken = default) {
            IReadOnlyList<string> requestedScopes = NormalizeScopes(scopes);
            ITokenAccess credential = _credentialFactory(requestedScopes);
            string accessToken = await credential
                .GetAccessTokenForRequestAsync(null, cancellationToken)
                .ConfigureAwait(false);

            if (string.IsNullOrWhiteSpace(accessToken)) {
                throw new InvalidOperationException("Google.Apis.Auth returned an empty access token.");
            }

            DateTimeOffset expiresAt = ResolveExpiry(credential, _fallbackTokenLifetime);
            IReadOnlyList<string> grantedScopes = ResolveGrantedScopes(credential, requestedScopes);
            return new GoogleWorkspaceAccessToken(accessToken, expiresAt, grantedScopes);
        }

        private static IReadOnlyList<string> NormalizeScopes(IEnumerable<string>? scopes) {
            return (scopes ?? Array.Empty<string>())
                .Where(scope => !string.IsNullOrWhiteSpace(scope))
                .Select(scope => scope.Trim())
                .Distinct(StringComparer.Ordinal)
                .OrderBy(scope => scope, StringComparer.Ordinal)
                .ToArray();
        }

        private static TimeSpan ValidateLifetime(TimeSpan? lifetime) {
            TimeSpan value = lifetime ?? DefaultLifetime;
            if (value <= TimeSpan.Zero) {
                throw new ArgumentOutOfRangeException(nameof(lifetime), "The fallback token lifetime must be positive.");
            }

            return value;
        }

        private static DateTimeOffset ResolveExpiry(ITokenAccess credential, TimeSpan fallbackLifetime) {
            UserCredential? userCredential = credential as UserCredential;
            if (credential is GoogleCredential googleCredential) {
                userCredential = googleCredential.UnderlyingCredential as UserCredential;
            }

            if (userCredential?.Token?.ExpiresInSeconds is long expiresInSeconds) {
                DateTime issuedUtc = DateTime.SpecifyKind(userCredential.Token.IssuedUtc, DateTimeKind.Utc);
                return new DateTimeOffset(issuedUtc).AddSeconds(expiresInSeconds);
            }

            return DateTimeOffset.UtcNow.Add(fallbackLifetime);
        }

        private static IReadOnlyList<string> ResolveGrantedScopes(
            ITokenAccess credential,
            IReadOnlyList<string> requestedScopes) {
            UserCredential? userCredential = credential as UserCredential;
            if (credential is GoogleCredential googleCredential) {
                userCredential = googleCredential.UnderlyingCredential as UserCredential;
            }

            string? granted = userCredential?.Token?.Scope;
            return string.IsNullOrWhiteSpace(granted)
                ? requestedScopes
                : granted!.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(scope => scope, StringComparer.Ordinal)
                    .ToArray();
        }
    }
}
