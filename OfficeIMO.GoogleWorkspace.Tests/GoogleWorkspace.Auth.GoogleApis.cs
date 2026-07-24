using Google.Apis.Auth.OAuth2;
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Auth.GoogleApis;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public class GoogleWorkspaceGoogleApisAuthTests {
        [Fact]
        public async Task Test_CredentialSource_NormalizesScopesAndUsesTokenAccess() {
            IReadOnlyList<string>? requestedScopes = null;
            var tokenAccess = new FakeTokenAccess((_, _) => Task.FromResult("access-token"));
            var source = new GoogleApisCredentialSource(tokenAccess, TimeSpan.FromMinutes(5));

            GoogleWorkspaceAccessToken token = await source.AcquireAccessTokenAsync(new[] {
                "scope-b",
                " scope-a ",
                "scope-b",
            });

            requestedScopes = token.Scopes;
            Assert.Equal("access-token", token.AccessToken);
            Assert.Equal(new[] { "scope-a", "scope-b" }, requestedScopes);
            Assert.InRange(token.ExpiresAt, DateTimeOffset.UtcNow.AddMinutes(4), DateTimeOffset.UtcNow.AddMinutes(6));
            Assert.Equal(1, tokenAccess.RequestCount);
        }

        [Fact]
        public async Task Test_DataStoreAdapter_DelegatesWithoutChoosingPlaintextStorage() {
            var store = new MemoryTokenStore();
            var adapter = new GoogleApisDataStoreAdapter(store);

            await adapter.StoreAsync("credential", new StoredToken { AccessToken = "secret" });
            StoredToken restored = await adapter.GetAsync<StoredToken>("credential");
            await adapter.DeleteAsync<StoredToken>("credential");

            Assert.Equal("secret", restored.AccessToken);
            Assert.Null(await adapter.GetAsync<StoredToken>("credential"));
        }

        [Fact]
        public async Task Test_InstalledApplicationAuthorization_RequiresExplicitTokenStoreBeforeInteraction() {
            var options = new GoogleInstalledApplicationAuthorizationOptions {
                ClientSecrets = new ClientSecrets { ClientId = "client-id", ClientSecret = "client-secret" },
                Scopes = new[] { GoogleWorkspaceScopeCatalog.DriveFile },
                UserId = "local-user",
            };

            InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(
                () => GoogleInstalledApplicationAuthorization.AuthorizeCredentialAsync(options));

            Assert.Contains("does not default OAuth refresh tokens to plaintext files", exception.Message);
        }

        [Fact]
        public async Task Test_InstalledApplicationAuthorization_RequiresExplicitStableUserId() {
            var options = new GoogleInstalledApplicationAuthorizationOptions {
                ClientSecrets = new ClientSecrets { ClientId = "client-id", ClientSecret = "client-secret" },
                Scopes = new[] { GoogleWorkspaceScopeCatalog.DriveFile },
                TokenStore = new MemoryTokenStore(),
            };

            InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(
                () => GoogleInstalledApplicationAuthorization.AuthorizeCredentialAsync(options));

            Assert.Contains("stable local user ID", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_CredentialSource_RejectsInvalidFallbackLifetime() {
            Assert.Throws<ArgumentOutOfRangeException>(
                () => new GoogleApisCredentialSource(new FakeTokenAccess((_, _) => Task.FromResult("unused")), TimeSpan.Zero));
        }

        private sealed class FakeTokenAccess : ITokenAccess {
            private readonly Func<string?, CancellationToken, Task<string>> _getToken;

            public FakeTokenAccess(Func<string?, CancellationToken, Task<string>> getToken) {
                _getToken = getToken;
            }

            public int RequestCount { get; private set; }

            public Task<string> GetAccessTokenForRequestAsync(
                string? authUri = null,
                CancellationToken cancellationToken = default) {
                RequestCount++;
                return _getToken(authUri, cancellationToken);
            }
        }

        private sealed class MemoryTokenStore : IGoogleWorkspaceTokenStore {
            private readonly ConcurrentDictionary<string, object?> _values = new ConcurrentDictionary<string, object?>();

            public Task StoreAsync<T>(string key, T value) {
                _values[key] = value;
                return Task.CompletedTask;
            }

            public Task DeleteAsync<T>(string key) {
                _values.TryRemove(key, out _);
                return Task.CompletedTask;
            }

            public Task<T?> GetAsync<T>(string key) {
                return Task.FromResult(_values.TryGetValue(key, out object? value) ? (T?)value : default);
            }

            public Task ClearAsync() {
                _values.Clear();
                return Task.CompletedTask;
            }
        }

        private sealed class StoredToken {
            public string? AccessToken { get; set; }
        }
    }
}
