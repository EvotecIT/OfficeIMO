using OfficeIMO.GoogleWorkspace;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public class GoogleServiceAccountCredentialSourceTests {
        private const string ServiceAccountJson = @"{
  ""client_email"": ""officeimo-tests@example.iam.gserviceaccount.com"",
  ""private_key"": ""-----BEGIN PRIVATE KEY-----\nMIICdwIBADANBgkqhkiG9w0BAQEFAASCAmEwggJdAgEAAoGBANF8RQeeF14651ue\nwmsDa5wo3CRRm9DKBxKml+CG+pGspQDKlXiuPLiFFyIw0rNLrRERTiYwnbzVfvN6\nsVwaUWtzQjO/h3+C1i8cZc+uw1/V3uQW5L/gwIGk3TtgkdJyJp7a2OpXoBehha2S\nsWpo234kTHfFta3vKU7f5f6+lsPhAgMBAAECgYBwefzGXkfFvHLEarWQp8F7kyTA\nC2FR9WdeyDv7vf2DgeMGTb97kHHh0PPe08ANrLA73cLMFoZbAXasXFAmV6smwQoA\nIPOc8RoK11Pmr3GRUr52Yhp4H+K/ZFIyxUoCQodvDkZY5cPO5a6/cMq94sxMgRat\n546Q3tN+mcPA2uAE0QJBANldsZP0/0f3VqTxWQCZCQlCgj+il3/FPKVgJwyZEV0p\nuTtEEVXQ+M1Ka7qgUVr+BrhogCycH0+o651LG5t4He8CQQD2uAE7Z64Hi0S2NJoK\n6GuLvyZEcEQunb6lcGXmwCgJbCrBJ1F0X7e2xPu24D9IDQW246eGgEA3HTsXhMfZ\nwAsvAkEAuAbP+iD5JDeufnTq0ku+T72kQjXop78YCjcuuEa7YbGaZifJuWrzyfKQ\n5G8oka3xiJzIr3v6MlokKIZXODfotwJBAOEI6s7Ffd4RsKXFCvCCGH5J5tyrzfT7\nGwaJo9i6Uoptp/2wIELWf5psx++BURcmEZ1EvuwWlPvwZJLKIQPDgC8CQG4v5XDl\nc131hSCWuXmn2fsXbv73UL9e1crFzPDwYSzlTYP9ya4UihF08b2geD5D5Uxa9L+y\nZTbjvDfKKybybrk=\n-----END PRIVATE KEY-----"",
  ""token_uri"": ""https://oauth.example.test/token""
}";

        [Fact]
        public async Task Test_GoogleServiceAccountCredentialSource_RequestsJwtBearerToken_AndCachesResponse() {
            if (!IsServiceAccountKeyImportSupported()) {
                return;
            }

            int requestCount = 0;
            string? assertion = null;

            var handler = new FakeHttpMessageHandler(async request => {
                requestCount++;
                string body = await request.Content!.ReadAsStringAsync().ConfigureAwait(false);
                var form = ParseFormUrlEncoded(body);
                assertion = form["assertion"];

                Assert.Equal("urn:ietf:params:oauth:grant-type:jwt-bearer", form["grant_type"]);

                return new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StringContent("{\"access_token\":\"service-account-token\",\"expires_in\":3600}", Encoding.UTF8, "application/json")
                };
            });

            var sessionOptions = new GoogleWorkspaceSessionOptions {
                HttpClient = new HttpClient(handler),
            };

            var source = GoogleServiceAccountCredentialSource.FromJson(ServiceAccountJson, sessionOptions);

            var first = await source.AcquireAccessTokenAsync(new[] {
                GoogleWorkspaceScopeCatalog.DriveFile,
                GoogleWorkspaceScopeCatalog.Documents,
            });
            var second = await source.AcquireAccessTokenAsync(new[] {
                GoogleWorkspaceScopeCatalog.Documents,
                GoogleWorkspaceScopeCatalog.DriveFile,
            });

            Assert.Equal("service-account-token", first.AccessToken);
            Assert.Equal("service-account-token", second.AccessToken);
            Assert.Equal(1, requestCount);
            Assert.NotNull(assertion);

            using JsonDocument header = JsonDocument.Parse(DecodeJwtPart(assertion!, 0));
            using JsonDocument payload = JsonDocument.Parse(DecodeJwtPart(assertion!, 1));

            Assert.Equal("RS256", header.RootElement.GetProperty("alg").GetString());
            Assert.Equal("JWT", header.RootElement.GetProperty("typ").GetString());
            Assert.Equal("officeimo-tests@example.iam.gserviceaccount.com", payload.RootElement.GetProperty("iss").GetString());
            Assert.Equal("https://oauth.example.test/token", payload.RootElement.GetProperty("aud").GetString());
            Assert.Equal(
                "https://www.googleapis.com/auth/documents https://www.googleapis.com/auth/drive.file",
                payload.RootElement.GetProperty("scope").GetString());
            Assert.False(payload.RootElement.TryGetProperty("sub", out _));
        }

        [Fact]
        public async Task Test_GoogleServiceAccountCredentialSource_IncludesDelegatedSubject_WhenConfigured() {
            if (!IsServiceAccountKeyImportSupported()) {
                return;
            }

            string? assertion = null;

            var handler = new FakeHttpMessageHandler(async request => {
                string body = await request.Content!.ReadAsStringAsync().ConfigureAwait(false);
                assertion = ParseFormUrlEncoded(body)["assertion"];

                return new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StringContent("{\"access_token\":\"delegated-token\",\"expires_in\":1800}", Encoding.UTF8, "application/json")
                };
            });

            var sessionOptions = new GoogleWorkspaceSessionOptions {
                HttpClient = new HttpClient(handler),
                SubjectUser = "admin@example.com",
                UseDomainWideDelegation = true,
            };

            var source = GoogleServiceAccountCredentialSource.FromJson(ServiceAccountJson, sessionOptions);
            var token = await source.AcquireAccessTokenAsync(GoogleWorkspaceScopeCatalog.SheetsAuthoring);

            Assert.Equal("delegated-token", token.AccessToken);
            Assert.NotNull(assertion);

            using JsonDocument payload = JsonDocument.Parse(DecodeJwtPart(assertion!, 1));
            Assert.Equal("admin@example.com", payload.RootElement.GetProperty("sub").GetString());
        }

        [Fact]
        public async Task Test_GoogleServiceAccountCredentialSource_FromFile_LoadsServiceAccountJson() {
            if (!IsServiceAccountKeyImportSupported()) {
                return;
            }

            string filePath = Path.Combine(Path.GetTempPath(), "officeimo-google-service-account-" + Guid.NewGuid().ToString("N") + ".json");

            try {
                File.WriteAllText(filePath, ServiceAccountJson);

                var handler = new FakeHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StringContent("{\"access_token\":\"file-token\",\"expires_in\":1200}", Encoding.UTF8, "application/json")
                }));

                var sessionOptions = new GoogleWorkspaceSessionOptions {
                    HttpClient = new HttpClient(handler),
                };

                var source = GoogleServiceAccountCredentialSource.FromFile(filePath, sessionOptions);
                var token = await source.AcquireAccessTokenAsync(new[] {
                    GoogleWorkspaceScopeCatalog.DriveFile,
                });

                Assert.Equal("file-token", token.AccessToken);
                Assert.Equal(new[] {
                    GoogleWorkspaceScopeCatalog.DriveFile,
                }, token.Scopes);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleServiceAccountCredentialSource_SurfacesOAuthErrorDetails_WhenTokenExchangeFails() {
            if (!IsServiceAccountKeyImportSupported()) {
                return;
            }

            var handler = new FakeHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest) {
                Content = new StringContent("{\"error\":\"unauthorized_client\",\"error_description\":\"Domain-wide delegation is not enabled for this service account.\"}", Encoding.UTF8, "application/json")
            }));

            var sessionOptions = new GoogleWorkspaceSessionOptions {
                HttpClient = new HttpClient(handler),
                SubjectUser = "admin@example.com",
                UseDomainWideDelegation = true,
            };

            var source = GoogleServiceAccountCredentialSource.FromJson(ServiceAccountJson, sessionOptions);
            var exception = await Assert.ThrowsAsync<HttpRequestException>(() =>
                source.AcquireAccessTokenAsync(GoogleWorkspaceScopeCatalog.DocsAuthoring));

            Assert.Contains("unauthorized_client", exception.Message, StringComparison.Ordinal);
            Assert.Contains("Domain-wide delegation is not enabled", exception.Message, StringComparison.Ordinal);
        }

        private static Dictionary<string, string> ParseFormUrlEncoded(string body) {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (string part in body.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries)) {
                string[] pieces = part.Split(new[] { '=' }, 2);
                string key = Uri.UnescapeDataString(pieces[0].Replace("+", "%20"));
                string value = pieces.Length > 1
                    ? Uri.UnescapeDataString(pieces[1].Replace("+", "%20"))
                    : string.Empty;

                result[key] = value;
            }

            return result;
        }

        private static string DecodeJwtPart(string jwt, int index) {
            string[] parts = jwt.Split('.');
            string base64 = parts[index]
                .Replace('-', '+')
                .Replace('_', '/');

            switch (base64.Length % 4) {
                case 2:
                    base64 += "==";
                    break;
                case 3:
                    base64 += "=";
                    break;
            }

            return Encoding.UTF8.GetString(Convert.FromBase64String(base64));
        }

        private static bool IsServiceAccountKeyImportSupported() {
#if NET472
            return false;
#else
            return true;
#endif
        }

        private sealed class FakeHttpMessageHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;

            public FakeHttpMessageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) {
                _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
                return _handler(request);
            }
        }
    }
}
