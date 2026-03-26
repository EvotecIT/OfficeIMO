using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Credential source that acquires Google access tokens by exchanging a signed service-account JWT assertion.
    /// </summary>
    public sealed class GoogleServiceAccountCredentialSource : IGoogleWorkspaceCredentialSource, IDisposable {
        private const string DefaultTokenEndpoint = "https://oauth2.googleapis.com/token";
        private static readonly TimeSpan DefaultTokenLifetime = TimeSpan.FromMinutes(55);
        private static readonly TimeSpan DefaultRefreshSkew = TimeSpan.FromMinutes(5);

        private readonly Dictionary<string, GoogleWorkspaceAccessToken> _tokenCache = new Dictionary<string, GoogleWorkspaceAccessToken>(StringComparer.Ordinal);
        private readonly SemaphoreSlim _cacheLock = new SemaphoreSlim(1, 1);
        private readonly HttpClient _httpClient;
        private readonly bool _disposeHttpClient;
        private readonly RSAParameters _privateKeyParameters;

        public GoogleServiceAccountCredentialSource(
            string clientEmail,
            string privateKeyPem,
            string? tokenEndpoint = null,
            GoogleWorkspaceSessionOptions? sessionOptions = null) {
            if (string.IsNullOrWhiteSpace(clientEmail)) throw new ArgumentException("Client email is required.", nameof(clientEmail));
            if (string.IsNullOrWhiteSpace(privateKeyPem)) throw new ArgumentException("Private key PEM is required.", nameof(privateKeyPem));

            ClientEmail = clientEmail;
            TokenEndpoint = string.IsNullOrWhiteSpace(tokenEndpoint) ? DefaultTokenEndpoint : tokenEndpoint!;
            SubjectUser = sessionOptions?.SubjectUser;
            UseDomainWideDelegation = sessionOptions?.UseDomainWideDelegation == true;
            TokenLifetime = DefaultTokenLifetime;
            RefreshSkew = DefaultRefreshSkew;

            if (sessionOptions?.HttpClient != null) {
                _httpClient = sessionOptions.HttpClient;
                _disposeHttpClient = false;
            } else {
                _httpClient = new HttpClient {
                    Timeout = sessionOptions?.RequestTimeout ?? TimeSpan.FromSeconds(100)
                };
                _disposeHttpClient = true;
            }

            _privateKeyParameters = GoogleServiceAccountPemKeyLoader.LoadRsaPrivateKey(privateKeyPem);
        }

        public string ClientEmail { get; }
        public string TokenEndpoint { get; }
        public string? SubjectUser { get; }
        public bool UseDomainWideDelegation { get; }
        public TimeSpan TokenLifetime { get; }
        public TimeSpan RefreshSkew { get; }

        public static GoogleServiceAccountCredentialSource FromJson(
            string serviceAccountJson,
            GoogleWorkspaceSessionOptions? sessionOptions = null) {
            if (string.IsNullOrWhiteSpace(serviceAccountJson)) throw new ArgumentException("Service account JSON is required.", nameof(serviceAccountJson));

            var payload = JsonSerializer.Deserialize<GoogleServiceAccountJsonPayload>(serviceAccountJson);
            if (payload == null) throw new InvalidOperationException("Service account JSON could not be parsed.");
            if (string.IsNullOrWhiteSpace(payload.ClientEmail)) throw new InvalidOperationException("Service account JSON is missing client_email.");
            if (string.IsNullOrWhiteSpace(payload.PrivateKey)) throw new InvalidOperationException("Service account JSON is missing private_key.");

            return new GoogleServiceAccountCredentialSource(
                payload.ClientEmail!,
                payload.PrivateKey!,
                payload.TokenUri,
                sessionOptions);
        }

        public static GoogleServiceAccountCredentialSource FromFile(
            string serviceAccountJsonPath,
            GoogleWorkspaceSessionOptions? sessionOptions = null) {
            if (string.IsNullOrWhiteSpace(serviceAccountJsonPath)) throw new ArgumentException("Service account JSON path is required.", nameof(serviceAccountJsonPath));
            return FromJson(File.ReadAllText(serviceAccountJsonPath), sessionOptions);
        }

        public async Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(
            IEnumerable<string> scopes,
            CancellationToken cancellationToken = default) {
            var normalizedScopes = NormalizeScopes(scopes);
            string cacheKey = string.Join(" ", normalizedScopes);
            DateTimeOffset now = DateTimeOffset.UtcNow;

            await _cacheLock.WaitAsync(cancellationToken).ConfigureAwait(false);
            try {
                if (_tokenCache.TryGetValue(cacheKey, out GoogleWorkspaceAccessToken? cachedToken)
                    && cachedToken.ExpiresAt > now.Add(RefreshSkew)) {
                    return cachedToken;
                }

                GoogleWorkspaceAccessToken refreshedToken = await RequestAccessTokenAsync(normalizedScopes, now, cancellationToken).ConfigureAwait(false);
                _tokenCache[cacheKey] = refreshedToken;
                return refreshedToken;
            } finally {
                _cacheLock.Release();
            }
        }

        public void Dispose() {
            if (_disposeHttpClient) {
                _httpClient.Dispose();
            }

            _cacheLock.Dispose();
        }

        private async Task<GoogleWorkspaceAccessToken> RequestAccessTokenAsync(
            IReadOnlyList<string> scopes,
            DateTimeOffset now,
            CancellationToken cancellationToken) {
            string assertion = CreateSignedJwtAssertion(scopes, now);
            using var request = new HttpRequestMessage(HttpMethod.Post, TokenEndpoint) {
                Content = new FormUrlEncodedContent(new Dictionary<string, string> {
                    ["grant_type"] = "urn:ietf:params:oauth:grant-type:jwt-bearer",
                    ["assertion"] = assertion,
                })
            };

            using HttpResponseMessage response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
            string responseBody = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (!response.IsSuccessStatusCode) {
                string formattedError = GoogleWorkspaceApiErrorFormatter.Format(responseBody) ?? responseBody;
                if (!string.IsNullOrWhiteSpace(formattedError)) {
                    throw new HttpRequestException($"Google token exchange failed with status code {(int)response.StatusCode}: {formattedError}");
                }

                throw new HttpRequestException($"Google token exchange failed with status code {(int)response.StatusCode}: {responseBody}");
            }

            var tokenResponse = JsonSerializer.Deserialize<GoogleOAuthTokenResponse>(responseBody);
            if (tokenResponse == null || string.IsNullOrWhiteSpace(tokenResponse.AccessToken)) {
                throw new InvalidOperationException("Google token response did not contain access_token.");
            }

            DateTimeOffset expiresAt = now.AddSeconds(tokenResponse.ExpiresIn > 0 ? tokenResponse.ExpiresIn : TokenLifetime.TotalSeconds);
            string accessToken = tokenResponse.AccessToken!;
            IReadOnlyList<string> grantedScopes = string.IsNullOrWhiteSpace(tokenResponse.Scope)
                ? scopes
                : tokenResponse.Scope!.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            return new GoogleWorkspaceAccessToken(accessToken, expiresAt, grantedScopes);
        }

        private string CreateSignedJwtAssertion(IReadOnlyList<string> scopes, DateTimeOffset now) {
            long issuedAt = now.ToUnixTimeSeconds();
            long expiresAt = now.Add(TokenLifetime).ToUnixTimeSeconds();

            string headerJson = "{\"alg\":\"RS256\",\"typ\":\"JWT\"}";
            var payload = new Dictionary<string, object?> {
                ["iss"] = ClientEmail,
                ["scope"] = string.Join(" ", scopes),
                ["aud"] = TokenEndpoint,
                ["iat"] = issuedAt,
                ["exp"] = expiresAt,
            };

            if (UseDomainWideDelegation && !string.IsNullOrWhiteSpace(SubjectUser)) {
                payload["sub"] = SubjectUser;
            }

            string payloadJson = JsonSerializer.Serialize(payload);

            string encodedHeader = Base64UrlEncode(Encoding.UTF8.GetBytes(headerJson));
            string encodedPayload = Base64UrlEncode(Encoding.UTF8.GetBytes(payloadJson));
            string signingInput = encodedHeader + "." + encodedPayload;

            using RSA rsa = RSA.Create();
            rsa.ImportParameters(_privateKeyParameters);
            byte[] signature = rsa.SignData(Encoding.ASCII.GetBytes(signingInput), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

            return signingInput + "." + Base64UrlEncode(signature);
        }

        private static IReadOnlyList<string> NormalizeScopes(IEnumerable<string> scopes) {
            if (scopes == null) throw new ArgumentNullException(nameof(scopes));

            var normalizedScopes = scopes
                .Where(scope => !string.IsNullOrWhiteSpace(scope))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(scope => scope, StringComparer.Ordinal)
                .ToArray();

            if (normalizedScopes.Length == 0) {
                throw new ArgumentException("At least one scope is required.", nameof(scopes));
            }

            return normalizedScopes;
        }

        private static string Base64UrlEncode(byte[] bytes) {
            return Convert.ToBase64String(bytes)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        private sealed class GoogleServiceAccountJsonPayload {
            [JsonPropertyName("client_email")]
            public string? ClientEmail { get; set; }

            [JsonPropertyName("private_key")]
            public string? PrivateKey { get; set; }

            [JsonPropertyName("token_uri")]
            public string? TokenUri { get; set; }
        }

        private sealed class GoogleOAuthTokenResponse {
            [JsonPropertyName("access_token")]
            public string? AccessToken { get; set; }

            [JsonPropertyName("expires_in")]
            public int ExpiresIn { get; set; }

            [JsonPropertyName("scope")]
            public string? Scope { get; set; }
        }
    }

    internal static class GoogleServiceAccountPemKeyLoader {
        internal static RSAParameters LoadRsaPrivateKey(string privateKeyPem) {
            if (string.IsNullOrWhiteSpace(privateKeyPem)) throw new ArgumentException("Private key PEM is required.", nameof(privateKeyPem));

#if NET8_0_OR_GREATER
            using RSA rsa = RSA.Create();
            rsa.ImportFromPem(privateKeyPem);
            return rsa.ExportParameters(true);
#else
            throw new PlatformNotSupportedException("GoogleServiceAccountCredentialSource requires a runtime with native PKCS#8 PEM import support. On this target, acquire tokens externally and use StaticAccessTokenCredentialSource or DelegateGoogleWorkspaceCredentialSource.");
#endif
        }
    }
}
