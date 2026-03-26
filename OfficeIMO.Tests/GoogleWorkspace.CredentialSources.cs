using OfficeIMO.GoogleWorkspace;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public class GoogleWorkspaceCredentialSourceTests {
        [Fact]
        public async Task Test_StaticAccessTokenCredentialSource_UsesRequestedScopes_WhenExplicitScopesAreNotConfigured() {
            var source = new StaticAccessTokenCredentialSource(
                "access-token",
                new DateTimeOffset(2026, 3, 25, 12, 0, 0, TimeSpan.Zero));

            var token = await source.AcquireAccessTokenAsync(new[] {
                GoogleWorkspaceScopeCatalog.DriveFile,
                GoogleWorkspaceScopeCatalog.Documents,
            });

            Assert.Equal("access-token", token.AccessToken);
            Assert.Equal(new DateTimeOffset(2026, 3, 25, 12, 0, 0, TimeSpan.Zero), token.ExpiresAt);
            Assert.Equal(new[] {
                GoogleWorkspaceScopeCatalog.DriveFile,
                GoogleWorkspaceScopeCatalog.Documents,
            }, token.Scopes);
        }

        [Fact]
        public async Task Test_StaticAccessTokenCredentialSource_PrefersConfiguredScopes_WhenProvided() {
            var source = new StaticAccessTokenCredentialSource(
                "access-token",
                scopes: new[] {
                    GoogleWorkspaceScopeCatalog.WorkspaceAuthoring[0],
                    GoogleWorkspaceScopeCatalog.WorkspaceAuthoring[1],
                });

            var token = await source.AcquireAccessTokenAsync(new[] {
                GoogleWorkspaceScopeCatalog.DriveReadonly,
            });

            Assert.Equal(new[] {
                GoogleWorkspaceScopeCatalog.WorkspaceAuthoring[0],
                GoogleWorkspaceScopeCatalog.WorkspaceAuthoring[1],
            }, token.Scopes);
        }

        [Fact]
        public async Task Test_DelegateGoogleWorkspaceCredentialSource_ForwardsScopesAndCancellationToken() {
            IReadOnlyList<string>? capturedScopes = null;
            CancellationToken capturedCancellationToken = default;

            using var cancellationTokenSource = new CancellationTokenSource();

            var source = new DelegateGoogleWorkspaceCredentialSource((scopes, cancellationToken) => {
                capturedScopes = scopes;
                capturedCancellationToken = cancellationToken;

                return Task.FromResult(new GoogleWorkspaceAccessToken(
                    "delegated-token",
                    DateTimeOffset.UtcNow.AddMinutes(10),
                    scopes));
            });

            var token = await source.AcquireAccessTokenAsync(new[] {
                GoogleWorkspaceScopeCatalog.DriveFile,
                GoogleWorkspaceScopeCatalog.Spreadsheets,
            }, cancellationTokenSource.Token);

            Assert.Equal("delegated-token", token.AccessToken);
            Assert.Equal(new[] {
                GoogleWorkspaceScopeCatalog.DriveFile,
                GoogleWorkspaceScopeCatalog.Spreadsheets,
            }, capturedScopes);
            Assert.Equal(cancellationTokenSource.Token, capturedCancellationToken);
        }
    }
}
