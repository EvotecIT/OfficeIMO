using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class GoogleWorkspaceRoadmapContractTests {
        [Fact]
        public async Task MutationRequiresExplicitPolicyAndEmitsReceipt() {
            using var http = new HttpClient(new Handler(_ => Json("{\"id\":\"folder-1\",\"version\":\"1\"}")));
            using (var unguarded = new GoogleDriveClient(new GoogleWorkspaceSession(
                new StaticAccessTokenCredentialSource("token"), new GoogleWorkspaceSessionOptions { HttpClient = http }))) {
                await Assert.ThrowsAsync<InvalidOperationException>(() => unguarded.CreateFolderAsync("blocked"));
            }

            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            GoogleWorkspaceSession session = Session(http, receipts);
            using var guarded = new GoogleDriveClient(session);
            GoogleDriveFile file = await guarded.CreateFolderAsync("allowed");

            Assert.Equal("folder-1", file.Id);
            GoogleWorkspaceOperationReceipt receipt = Assert.Single(receipts);
            Assert.True(receipt.Succeeded);
            Assert.Equal("test@example.com", receipt.Policy.Account);
            Assert.NotEmpty(receipt.Policy.Scopes);
            Assert.Equal(GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision, receipt.Policy.ExpectedRevision);
            Assert.Equal(GoogleWorkspaceDataLossDecision.RejectPotentialLoss, receipt.Policy.DataLossDecision);
            Assert.Equal(GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate, receipt.RevisionPreconditionKind);
            Assert.Null(receipt.EnforcedRevision);
        }

        [Fact]
        public async Task DriveMutationUsesTheScopeSnapshotAcquiredByTheClient() {
            using var http = new HttpClient(new Handler(_ => Json("{\"id\":\"folder-1\",\"version\":\"1\"}")));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            var configuredScopes = new List<string> { GoogleWorkspaceScopeCatalog.DriveFile };
            var driveOptions = new GoogleDriveClientOptions { WriteScopes = configuredScopes };
            using var client = new GoogleDriveClient(Session(http, receipts), driveOptions);
            configuredScopes[0] = GoogleWorkspaceScopeCatalog.Drive;

            await client.CreateFolderAsync("scope-snapshot");

            GoogleWorkspaceOperationReceipt receipt = Assert.Single(receipts);
            Assert.Equal(new[] { GoogleWorkspaceScopeCatalog.DriveFile }, receipt.Policy.Scopes);
        }

        [Theory]
        [InlineData("other@example.com", false)]
        [InlineData(null, false)]
        [InlineData("test@example.com", true)]
        public async Task SessionRejectsCredentialIdentityOrScopesThatDoNotMatchTheMutationContract(
            string? credentialAccount, bool useWrongScopes) {
            var options = new GoogleWorkspaceSessionOptions { ExpectedAccount = "test@example.com" };
            IReadOnlyList<string> tokenScopes = useWrongScopes
                ? new[] { GoogleWorkspaceScopeCatalog.Drive }
                : new[] { GoogleWorkspaceScopeCatalog.DriveFile };
            IGoogleWorkspaceCredentialSource credentialSource = credentialAccount == null
                ? new StaticAccessTokenCredentialSource("token", scopes: tokenScopes)
                : new DelegateGoogleWorkspaceCredentialSource((_, _) => Task.FromResult(
                    GoogleWorkspaceAccessToken.FromVerifiedCredential("token", DateTimeOffset.UtcNow.AddMinutes(5),
                        new GoogleWorkspaceCredentialBinding(credentialAccount, tokenScopes))));
            var session = new GoogleWorkspaceSession(credentialSource, options);

            await Assert.ThrowsAsync<InvalidOperationException>(() => session.AcquireAccessTokenAsync(
                new[] { GoogleWorkspaceScopeCatalog.DriveFile }));
        }

        [Fact]
        public async Task SessionRejectsMatchingCallerLabelWithoutProviderEvidence() {
            var options = new GoogleWorkspaceSessionOptions { ExpectedAccount = "test@example.com" };
            var session = new GoogleWorkspaceSession(new StaticAccessTokenCredentialSource(
                "token", null, new[] { GoogleWorkspaceScopeCatalog.DriveFile }, "test@example.com"), options);

            InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                session.AcquireAccessTokenAsync(new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Contains("provider-verified", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task OptionsOnlyTransportCannotBypassSessionCredentialVerification() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            GoogleWorkspaceSessionOptions options = MutationOptions(http, _ => { });
            using var transport = new GoogleWorkspaceHttpTransport(options);

            InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                transport.SendJsonAsync<object>("token", HttpMethod.Post,
                    "https://www.googleapis.com/drive/v3/files", new { name = "blocked" },
                    GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API",
                    new TranslationReport(), mutationKind: GoogleWorkspaceMutationKind.Create,
                    requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Contains("transport bound", exception.Message, StringComparison.Ordinal);
            Assert.Equal(0, requests);
        }

        [Fact]
        public async Task MutationAppliesExpectedEntityTagAsIfMatch() {
            string? ifMatch = null;
            using var http = new HttpClient(new Handler(request => {
                ifMatch = request.Headers.IfMatch.Single().Tag;
                return Json("{}");
            }));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                OperationReceiptSink = receipts.Add,
            };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, context.RequiredScopes, context.Target,
                "\"revision-7\"", context.MaxRetryCount, context.MaxRetryElapsedTime,
                context.RateLimitPolicy, GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile);

            await transport.SendJsonAsync<object>("token", new HttpMethod("PATCH"),
                "https://www.googleapis.com/drive/v3/files/file-1", new { name = "updated" },
                GoogleWorkspaceRequestSafety.Idempotent, "Google Drive API", new TranslationReport(),
                requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile });

            Assert.Equal("\"revision-7\"", ifMatch);
            GoogleWorkspaceOperationReceipt receipt = Assert.Single(receipts);
            Assert.True(receipt.Succeeded);
            Assert.Equal(GoogleWorkspaceRevisionPreconditionKind.HttpEntityTag, receipt.RevisionPreconditionKind);
            Assert.Equal("\"revision-7\"", receipt.EnforcedRevision);
        }

        [Theory]
        [InlineData("drive-version:7")]
        [InlineData("*")]
        [InlineData("W/\"revision-7\"")]
        public async Task MutationRejectsObservedRevisionThatCannotBeEnforced(string expectedRevision) {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                OperationReceiptSink = _ => { },
            };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, context.RequiredScopes, context.Target,
                expectedRevision, context.MaxRetryCount, context.MaxRetryElapsedTime,
                context.RateLimitPolicy, GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile);

            await Assert.ThrowsAsync<InvalidOperationException>(() => transport.SendJsonAsync<object>(
                "token", new HttpMethod("PATCH"), "https://www.googleapis.com/drive/v3/files/file-1",
                new { name = "updated" }, GoogleWorkspaceRequestSafety.Idempotent,
                "Google Drive API", new TranslationReport(),
                requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Equal(0, requests);
        }

        [Fact]
        public async Task ResourceAbsentRevisionCannotAuthorizePostUpdate() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            var options = MutationOptions(http, _ => { });
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.Documents);

            await Assert.ThrowsAsync<InvalidOperationException>(() => transport.SendJsonAsync<object>(
                "token", HttpMethod.Post, "https://docs.googleapis.com/v1/documents/doc-1:batchUpdate",
                new { requests = Array.Empty<object>() }, GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Docs API", new TranslationReport(),
                mutationKind: GoogleWorkspaceMutationKind.Update,
                revisionPrecondition: GoogleWorkspaceRevisionPrecondition.PayloadRevision("docs-revision-7"),
                requiredScopes: new[] { GoogleWorkspaceScopeCatalog.Documents }));

            Assert.Equal(0, requests);
        }

        [Fact]
        public async Task PayloadRevisionMustMatchPolicyAndIsRecordedAsEnforced() {
            bool hadIfMatch = true;
            using var http = new HttpClient(new Handler(request => {
                hadIfMatch = request.Headers.IfMatch.Any();
                return Json("{}");
            }));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                OperationReceiptSink = receipts.Add,
            };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, context.RequiredScopes, context.Target,
                context.AdapterExpectedRevision!, context.MaxRetryCount, context.MaxRetryElapsedTime,
                context.RateLimitPolicy, GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.Documents);

            await transport.SendJsonAsync<object>("token", HttpMethod.Post,
                "https://docs.googleapis.com/v1/documents/doc-1:batchUpdate",
                new { requests = Array.Empty<object>() }, GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Docs API", new TranslationReport(), mutationKind: GoogleWorkspaceMutationKind.Update,
                revisionPrecondition: GoogleWorkspaceRevisionPrecondition.PayloadRevision("docs-revision-7"),
                requiredScopes: new[] { GoogleWorkspaceScopeCatalog.Documents });

            Assert.False(hadIfMatch);
            GoogleWorkspaceOperationReceipt receipt = Assert.Single(receipts);
            Assert.Equal(GoogleWorkspaceRevisionPreconditionKind.PayloadRevision, receipt.RevisionPreconditionKind);
            Assert.Equal("docs-revision-7", receipt.EnforcedRevision);
            Assert.Equal(receipt.Policy.ExpectedRevision, receipt.EnforcedRevision);
        }

        [Fact]
        public async Task ReceiptSinkFailureKeepsRemoteSuccessDistinguishable() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            var options = MutationOptions(http, _ => throw new IOException("receipt store unavailable"));
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile);

            GoogleWorkspaceReceiptPersistenceException exception =
                await Assert.ThrowsAsync<GoogleWorkspaceReceiptPersistenceException>(() =>
                    transport.SendJsonAsync<object>("token", HttpMethod.Post,
                        "https://www.googleapis.com/drive/v3/files", new { name = "created" },
                        GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API", new TranslationReport(),
                        mutationKind: GoogleWorkspaceMutationKind.Create,
                        requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Equal(1, requests);
            Assert.True(exception.RemoteOperationSucceeded);
            Assert.True(exception.Receipt.Succeeded);
        }

        [Fact]
        public async Task MalformedMutationResponseStillRecordsRemoteSuccess() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => {
                requests++;
                return new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StringContent("{", Encoding.UTF8, "application/json"),
                };
            }));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            using var transport = MutationTransport(MutationOptions(http, receipts.Add),
                GoogleWorkspaceScopeCatalog.DriveFile);

            await Assert.ThrowsAsync<System.Text.Json.JsonException>(() =>
                transport.SendJsonAsync<GoogleDriveFile>("token", HttpMethod.Post,
                    "https://www.googleapis.com/drive/v3/files", new { name = "created" },
                    GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API", new TranslationReport(),
                    mutationKind: GoogleWorkspaceMutationKind.Create,
                    requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Equal(1, requests);
            GoogleWorkspaceOperationReceipt receipt = Assert.Single(receipts);
            Assert.True(receipt.Succeeded);
            Assert.Equal("completed", receipt.Outcome);
        }

        [Fact]
        public async Task IdempotentMutationDoesNotRetryAfterAcceptedResponseBodyFailure() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => {
                requests++;
                return new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ThrowingReadContent(),
                };
            }));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            GoogleWorkspaceSessionOptions options = MutationOptions(http, receipts.Add);
            options.MaxRetryCount = 2;
            options.RetryBaseDelay = TimeSpan.FromMilliseconds(1);
            options.RetryMaxDelay = TimeSpan.FromMilliseconds(1);
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, context.RequiredScopes, context.Target,
                "\"revision-1\"", context.MaxRetryCount, context.MaxRetryElapsedTime,
                context.RateLimitPolicy, GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile);

            await Assert.ThrowsAsync<IOException>(() => transport.SendJsonAsync<object>(
                "token", new HttpMethod("PATCH"),
                "https://www.googleapis.com/drive/v3/files/file-1", new { name = "committed" },
                GoogleWorkspaceRequestSafety.Idempotent, "Google Drive API", new TranslationReport(),
                requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Equal(1, requests);
            Assert.True(Assert.Single(receipts).Succeeded);
        }

        [Fact]
        public async Task RawIdempotentMutationDoesNotRetryAfterAcceptedResponseBodyFailure() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => {
                requests++;
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ThrowingReadContent() };
            }));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            GoogleWorkspaceSessionOptions options = MutationOptions(http, receipts.Add);
            options.MaxRetryCount = 2;
            options.RetryBaseDelay = TimeSpan.FromMilliseconds(1);
            options.RetryMaxDelay = TimeSpan.FromMilliseconds(1);
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, context.RequiredScopes, context.Target,
                "\"revision-1\"", context.MaxRetryCount, context.MaxRetryElapsedTime,
                context.RateLimitPolicy, GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile);

            await Assert.ThrowsAsync<IOException>(() => transport.SendRawAsync(
                "token", new HttpMethod("PATCH"),
                "https://www.googleapis.com/drive/v3/files/file-1", null,
                GoogleWorkspaceRequestSafety.Idempotent, "Google Drive API", new TranslationReport(),
                requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Equal(1, requests);
            Assert.True(Assert.Single(receipts).Succeeded);
        }

        [Theory]
        [InlineData("json", "network")]
        [InlineData("json", "timeout")]
        [InlineData("json", "canceled")]
        [InlineData("bytes", "network")]
        [InlineData("bytes", "timeout")]
        [InlineData("bytes", "canceled")]
        [InlineData("raw", "network")]
        [InlineData("raw", "timeout")]
        [InlineData("raw", "canceled")]
        public async Task NonIdempotentMutationWithoutResponseRequiresReconciliation(
            string responsePath, string failureKind) {
            int requests = 0;
            Exception transportFailure = failureKind switch {
                "timeout" => new TimeoutException("request timed out before response headers"),
                "canceled" => new TaskCanceledException("request canceled before response headers"),
                _ => new HttpRequestException("connection closed before response headers"),
            };
            using var http = new HttpClient(new Handler(_ => {
                requests++;
                return Task.FromException<HttpResponseMessage>(transportFailure);
            }));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            GoogleWorkspaceSessionOptions options = MutationOptions(http, receipts.Add);
            options.MaxRetryCount = 2;
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile);

            Task ExecuteAsync() {
                const string uri = "https://www.googleapis.com/drive/v3/files";
                var report = new TranslationReport();
                string[] scopes = { GoogleWorkspaceScopeCatalog.DriveFile };
                switch (responsePath) {
                    case "json":
                        return transport.SendJsonAsync<object>("token", HttpMethod.Post, uri,
                            new { name = "possibly-created" }, GoogleWorkspaceRequestSafety.NonIdempotent,
                            "Google Drive API", report, mutationKind: GoogleWorkspaceMutationKind.Create,
                            requiredScopes: scopes);
                    case "bytes":
                        return transport.SendBytesAsync("token", HttpMethod.Post, uri,
                            GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API", report,
                            mutationKind: GoogleWorkspaceMutationKind.Create, requiredScopes: scopes);
                    default:
                        return transport.SendRawAsync("token", HttpMethod.Post, uri, null,
                            GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API", report,
                            mutationKind: GoogleWorkspaceMutationKind.Create, requiredScopes: scopes);
                }
            }

            GoogleWorkspaceAmbiguousMutationException exception =
                await Assert.ThrowsAsync<GoogleWorkspaceAmbiguousMutationException>(ExecuteAsync);

            Assert.Equal(1, requests);
            Assert.Same(transportFailure, exception.InnerException);
            GoogleWorkspaceOperationReceipt receipt = Assert.Single(receipts);
            Assert.Same(receipt, exception.Receipt);
            Assert.False(receipt.Succeeded);
            Assert.True(receipt.IsOutcomeAmbiguous);
            Assert.Equal("ambiguous-no-response", receipt.Outcome);
            Assert.Equal(0, receipt.RetryCount);
        }

        [Theory]
        [InlineData("POST")]
        [InlineData("PATCH")]
        [InlineData("DELETE")]
        public async Task SafeClassificationCannotBypassMutationGuards(string method) {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            using var transport = new GoogleWorkspaceHttpTransport(new GoogleWorkspaceSessionOptions { HttpClient = http });

            await Assert.ThrowsAsync<InvalidOperationException>(() => transport.SendRawAsync(
                "token", new HttpMethod(method), "https://www.googleapis.com/drive/v3/files/file-1", null,
                GoogleWorkspaceRequestSafety.Safe, "Google Drive API", new TranslationReport(),
                mutationKind: GoogleWorkspaceMutationKind.Update,
                requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Equal(0, requests);
        }

        [Fact]
        public void LegacyCredentialConstructorsRemainAvailable() {
            Assert.NotNull(typeof(GoogleWorkspaceAccessToken).GetConstructor(new[] {
                typeof(string), typeof(DateTimeOffset), typeof(IReadOnlyList<string>)
            }));
            Assert.NotNull(typeof(StaticAccessTokenCredentialSource).GetConstructor(new[] {
                typeof(string), typeof(DateTimeOffset?), typeof(IReadOnlyList<string>)
            }));
        }

        [Fact]
        public async Task ReceiptSinkFailureDoesNotReplaceRemoteFailure() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => {
                requests++;
                return new HttpResponseMessage(HttpStatusCode.BadRequest) {
                    Content = new StringContent("{\"error\":{\"message\":\"bad request\"}}", Encoding.UTF8, "application/json")
                };
            }));
            var options = MutationOptions(http, _ => throw new IOException("receipt store unavailable"));
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile);

            GoogleWorkspaceApiException exception = await Assert.ThrowsAsync<GoogleWorkspaceApiException>(() =>
                transport.SendJsonAsync<object>("token", HttpMethod.Post,
                    "https://www.googleapis.com/drive/v3/files", new { name = "created" },
                    GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API", new TranslationReport(),
                    mutationKind: GoogleWorkspaceMutationKind.Create,
                    requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Equal(1, requests);
            var receiptFailure = Assert.IsType<GoogleWorkspaceReceiptPersistenceException>(
                exception.Data[GoogleWorkspaceReceiptPersistenceException.ExceptionDataKey]);
            Assert.False(receiptFailure.RemoteOperationSucceeded);
            Assert.False(receiptFailure.Receipt.Succeeded);
        }

        [Fact]
        public async Task MutationRejectsPolicyWhoseScopesDoNotMatchAdapterRequest() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                OperationReceiptSink = _ => { },
            };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, new[] { GoogleWorkspaceScopeCatalog.DriveFile }, context.Target,
                GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
                context.MaxRetryCount, context.MaxRetryElapsedTime, context.RateLimitPolicy,
                GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile,
                GoogleWorkspaceScopeCatalog.Documents);

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                transport.SendJsonAsync<object>("token", HttpMethod.Post,
                    "https://www.googleapis.com/drive/v3/files", new { name = "created" },
                    GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API",
                    new TranslationReport(), mutationKind: GoogleWorkspaceMutationKind.Create,
                    requiredScopes: new[] {
                        GoogleWorkspaceScopeCatalog.DriveFile,
                        GoogleWorkspaceScopeCatalog.Documents,
                    }));

            Assert.Equal(0, requests);
        }

        [Fact]
        public async Task MutationPolicyValidationUsesSingleRetrySnapshot() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                MaxRetryCount = 1,
                MaxRetryElapsedTime = TimeSpan.FromSeconds(30),
                OperationReceiptSink = _ => { },
            };
            options.OperationPolicyProvider = context => {
                options.MaxRetryCount = 9;
                options.MaxRetryElapsedTime = TimeSpan.FromMinutes(9);
                options.RateLimitPolicy = GoogleWorkspaceRateLimitPolicy.FailFast;
                return new GoogleWorkspaceOperationPolicy(
                    options.ExpectedAccount!, context.RequiredScopes, context.Target,
                    GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
                    options.MaxRetryCount, options.MaxRetryElapsedTime, options.RateLimitPolicy,
                    GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            };
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile);

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                transport.SendJsonAsync<object>("token", HttpMethod.Post,
                    "https://www.googleapis.com/drive/v3/files", new { name = "created" },
                    GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API",
                    new TranslationReport(), mutationKind: GoogleWorkspaceMutationKind.Create,
                    requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Equal(0, requests);
        }

        [Fact]
        public async Task MutationExecutionUsesTheRetrySnapshotExposedToPolicy() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => {
                requests++;
                return new HttpResponseMessage(HttpStatusCode.ServiceUnavailable) {
                    Content = new StringContent("{\"error\":{\"message\":\"retryable\"}}", Encoding.UTF8, "application/json")
                };
            }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                MaxRetryCount = 0,
                MaxRetryElapsedTime = TimeSpan.FromSeconds(30),
                OperationReceiptSink = _ => { },
            };
            options.OperationPolicyProvider = context => {
                options.MaxRetryCount = 9;
                options.MaxRetryElapsedTime = TimeSpan.FromMinutes(9);
                options.RateLimitPolicy = GoogleWorkspaceRateLimitPolicy.FailFast;
                return new GoogleWorkspaceOperationPolicy(
                    options.ExpectedAccount!, context.RequiredScopes, context.Target,
                    GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
                    context.MaxRetryCount, context.MaxRetryElapsedTime, context.RateLimitPolicy,
                    GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            };
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile);

            await Assert.ThrowsAsync<GoogleWorkspaceApiException>(() =>
                transport.SendJsonAsync<object>("token", HttpMethod.Post,
                    "https://www.googleapis.com/drive/v3/files", new { name = "created" },
                    GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API",
                    new TranslationReport(), mutationKind: GoogleWorkspaceMutationKind.Create,
                    requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile }));

            Assert.Equal(1, requests);
        }

        [Fact]
        public async Task Google403RateLimitReasonUsesConfiguredRetryPolicy() {
            int attempts = 0;
            using var http = new HttpClient(new Handler(_ => {
                attempts++;
                return attempts == 1
                    ? new HttpResponseMessage(HttpStatusCode.Forbidden) {
                        Content = new StringContent("{\"error\":{\"errors\":[{\"reason\":\"userRateLimitExceeded\"}]}}", Encoding.UTF8, "application/json")
                    }
                    : Json("{\"id\":\"file-1\"}");
            }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                MaxRetryCount = 1,
                RetryBaseDelay = TimeSpan.FromMilliseconds(1),
                RetryMaxDelay = TimeSpan.FromMilliseconds(1),
                RateLimitPolicy = GoogleWorkspaceRateLimitPolicy.HonorRetryAfter,
            };
            using var transport = new GoogleWorkspaceHttpTransport(options);

            await transport.SendJsonAsync<object>("token", HttpMethod.Get,
                "https://www.googleapis.com/drive/v3/files/file-1", null,
                GoogleWorkspaceRequestSafety.Safe, "Google Drive API", new TranslationReport());

            Assert.Equal(2, attempts);
        }

        [Theory]
        [InlineData("insufficientFilePermissions", GoogleWorkspaceRateLimitPolicy.HonorRetryAfter)]
        [InlineData("userRateLimitExceeded", GoogleWorkspaceRateLimitPolicy.FailFast)]
        public async Task Google403DoesNotRetryAuthorizationErrorsOrFailFastQuota(string reason,
            GoogleWorkspaceRateLimitPolicy rateLimitPolicy) {
            int attempts = 0;
            using var http = new HttpClient(new Handler(_ => {
                attempts++;
                return new HttpResponseMessage(HttpStatusCode.Forbidden) {
                    Content = new StringContent("{\"error\":{\"errors\":[{\"reason\":\"" + reason + "\"}]}}",
                        Encoding.UTF8, "application/json")
                };
            }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                MaxRetryCount = 1,
                RetryBaseDelay = TimeSpan.FromMilliseconds(1),
                RetryMaxDelay = TimeSpan.FromMilliseconds(1),
                RateLimitPolicy = rateLimitPolicy,
            };
            using var transport = new GoogleWorkspaceHttpTransport(options);

            await Assert.ThrowsAsync<GoogleWorkspaceApiException>(() => transport.SendJsonAsync<object>(
                "token", HttpMethod.Get, "https://www.googleapis.com/drive/v3/files/file-1", null,
                GoogleWorkspaceRequestSafety.Safe, "Google Drive API", new TranslationReport()));

            Assert.Equal(1, attempts);
        }

        [Fact]
        public async Task RetryResponseIsDisposedBeforeElapsedBudgetRejection() {
            var response = new TrackingResponseMessage(HttpStatusCode.ServiceUnavailable) {
                Content = new StringContent("retry later")
            };
            response.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(5));
            using var http = new HttpClient(new Handler(_ => response));
            var retryOptions = new GoogleWorkspaceRetryOptions(1, TimeSpan.FromSeconds(1),
                TimeSpan.FromSeconds(5), new GoogleWorkspaceSessionOptions {
                    MaxRetryCount = 1,
                    MaxRetryElapsedTime = TimeSpan.FromMilliseconds(10),
                    RetryBaseDelay = TimeSpan.FromSeconds(1),
                    RetryMaxDelay = TimeSpan.FromSeconds(5),
                });

            await Assert.ThrowsAsync<TimeoutException>(() => GoogleWorkspaceRetryPolicy.SendAsync(
                http, () => new HttpRequestMessage(HttpMethod.Get, "https://www.googleapis.com/drive/v3/files/file-1"),
                retryOptions, GoogleWorkspaceRequestSafety.Safe, CancellationToken.None));

            Assert.True(response.WasDisposed);
        }

        [Fact]
        public async Task Google403BodyReadFailureDisposesAndRetriesSafeRequest() {
            int attempts = 0;
            var firstResponse = new TrackingResponseMessage(HttpStatusCode.Forbidden) {
                Content = new ThrowingReadContent()
            };
            using var http = new HttpClient(new Handler(_ => {
                attempts++;
                return attempts == 1 ? firstResponse : Json("{\"id\":\"file-1\"}");
            }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                MaxRetryCount = 1,
                RetryBaseDelay = TimeSpan.FromMilliseconds(1),
                RetryMaxDelay = TimeSpan.FromMilliseconds(1),
            };
            using var transport = new GoogleWorkspaceHttpTransport(options);

            await transport.SendJsonAsync<object>("token", HttpMethod.Get,
                "https://www.googleapis.com/drive/v3/files/file-1", null,
                GoogleWorkspaceRequestSafety.Safe, "Google Drive API", new TranslationReport());

            Assert.Equal(2, attempts);
            Assert.True(firstResponse.WasDisposed);
        }

        [Fact]
        public async Task FinalResponseIsDisposedWhenElapsedBudgetExpiresBeforeReturn() {
            var response = new TrackingResponseMessage(HttpStatusCode.OK) {
                Content = new StringContent("{}", Encoding.UTF8, "application/json")
            };
            using var http = new HttpClient(new Handler(async _ => {
                await Task.Delay(TimeSpan.FromMilliseconds(30)).ConfigureAwait(false);
                return response;
            }));
            var retryOptions = new GoogleWorkspaceRetryOptions(0, TimeSpan.FromMilliseconds(1),
                TimeSpan.FromMilliseconds(1), new GoogleWorkspaceSessionOptions {
                    MaxRetryCount = 0,
                    MaxRetryElapsedTime = TimeSpan.FromMilliseconds(5),
                });

            await Assert.ThrowsAsync<TimeoutException>(() => GoogleWorkspaceRetryPolicy.SendAsync(
                http, () => new HttpRequestMessage(HttpMethod.Get, "https://www.googleapis.com/drive/v3/files/file-1"),
                retryOptions, GoogleWorkspaceRequestSafety.Safe, CancellationToken.None));

            Assert.True(response.WasDisposed);
        }

        [Fact]
        public async Task DeleteRequiresExplicitAcceptanceOfDataLoss() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            using (var refusing = new GoogleDriveClient(Session(http, new List<GoogleWorkspaceOperationReceipt>()))) {
                await Assert.ThrowsAsync<InvalidOperationException>(() => refusing.DeleteFileAsync("file-1"));
            }
            Assert.Equal(0, requests);

            using var accepting = new GoogleDriveClient(Session(http,
                new List<GoogleWorkspaceOperationReceipt>(), GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss,
                "delete file-1"));
            await accepting.DeleteFileAsync("file-1");
            Assert.Equal(1, requests);
        }

        [Fact]
        public async Task AdapterDeclaredDestructiveUpdateRequiresAcceptedLossAndRecordsIt() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            GoogleWorkspaceOperationContext? observed = null;
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                OperationReceiptSink = receipts.Add,
            };
            options.OperationPolicyProvider = context => {
                observed = context;
                return new GoogleWorkspaceOperationPolicy(
                    options.ExpectedAccount!, context.RequiredScopes, context.Target,
                    context.AdapterExpectedRevision!, context.MaxRetryCount,
                    context.MaxRetryElapsedTime, context.RateLimitPolicy,
                    GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            };
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.Documents);

            await Assert.ThrowsAsync<InvalidOperationException>(() => transport.SendJsonAsync<object>(
                "token", HttpMethod.Post,
                "https://docs.googleapis.com/v1/documents/doc-1:batchUpdate",
                new { requests = Array.Empty<object>() }, GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Docs API", new TranslationReport(),
                mutationKind: GoogleWorkspaceMutationKind.Update,
                revisionPrecondition: GoogleWorkspaceRevisionPrecondition.PayloadRevision("revision-7"),
                requiredScopes: new[] { GoogleWorkspaceScopeCatalog.Documents },
                potentialDataLoss: true));

            Assert.Equal(0, requests);
            Assert.True(Assert.IsType<GoogleWorkspaceOperationContext>(observed).PotentialDataLoss);

            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, context.RequiredScopes, context.Target,
                context.AdapterExpectedRevision!, context.MaxRetryCount,
                context.MaxRetryElapsedTime, context.RateLimitPolicy,
                GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss, "replace document content");
            await transport.SendJsonAsync<object>(
                "token", HttpMethod.Post,
                "https://docs.googleapis.com/v1/documents/doc-1:batchUpdate",
                new { requests = Array.Empty<object>() }, GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Docs API", new TranslationReport(),
                mutationKind: GoogleWorkspaceMutationKind.Update,
                revisionPrecondition: GoogleWorkspaceRevisionPrecondition.PayloadRevision("revision-7"),
                requiredScopes: new[] { GoogleWorkspaceScopeCatalog.Documents },
                potentialDataLoss: true);

            Assert.Equal(1, requests);
            Assert.Equal(GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss,
                Assert.Single(receipts).Policy.DataLossDecision);
        }

        [Fact]
        public async Task MutationTargetsExcludeSensitiveQueryValues() {
            const string secretMessage = "private-recipient-message";
            const string secretQuotaUser = "private-quota-user";
            GoogleWorkspaceOperationContext? observed = null;
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            using var http = new HttpClient(new Handler(_ => Json("{}")));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                QuotaUser = secretQuotaUser,
                OperationReceiptSink = receipts.Add,
            };
            options.OperationPolicyProvider = context => {
                observed = context;
                return new GoogleWorkspaceOperationPolicy(
                    options.ExpectedAccount!, context.RequiredScopes, context.Target,
                    GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
                    context.MaxRetryCount, context.MaxRetryElapsedTime, context.RateLimitPolicy,
                    GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            };
            using var transport = MutationTransport(options, GoogleWorkspaceScopeCatalog.DriveFile);

            await transport.SendJsonAsync<object>("token", HttpMethod.Post,
                "https://www.googleapis.com/drive/v3/files/file-1/permissions?emailMessage=" + secretMessage,
                new { role = "reader", type = "user" }, GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Drive API", new TranslationReport(),
                mutationKind: GoogleWorkspaceMutationKind.Create,
                requiredScopes: new[] { GoogleWorkspaceScopeCatalog.DriveFile });

            string target = Assert.IsType<GoogleWorkspaceOperationContext>(observed).Target;
            Assert.Equal("https://www.googleapis.com/drive/v3/files/file-1/permissions", target);
            GoogleWorkspaceOperationReceipt receipt = Assert.Single(receipts);
            Assert.Equal(target, receipt.Target);
            Assert.DoesNotContain(secretMessage, receipt.Target, StringComparison.Ordinal);
            Assert.DoesNotContain(secretQuotaUser, receipt.Target, StringComparison.Ordinal);
        }

        [Fact]
        public async Task ResumableUploadCheckpointSurvivesRestartAndReconcilesServerOffset() {
            const int total = 300 * 1024;
            bool firstChunkCommitted = false;
            using var http = new HttpClient(new Handler(async request => {
                if (request.Method == HttpMethod.Post) { var response = Json("{}"); response.Headers.Location = new Uri("https://upload.googleapis.com/durable"); return response; }
                string range = request.Content!.Headers.GetValues("Content-Range").Single();
                if (range == $"bytes */{total}") {
                    var status = new HttpResponseMessage((HttpStatusCode)308) { Content = new StringContent(string.Empty) };
                    if (firstChunkCommitted) status.Headers.TryAddWithoutValidation("Range", "bytes=0-262143");
                    return status;
                }
                _ = await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                if (range.StartsWith("bytes 0-", StringComparison.Ordinal)) {
                    firstChunkCommitted = true; var accepted = new HttpResponseMessage((HttpStatusCode)308) { Content = new StringContent(string.Empty) };
                    accepted.Headers.TryAddWithoutValidation("Range", "bytes=0-262143"); return accepted;
                }
                return new HttpResponseMessage(HttpStatusCode.Created) { Content = new StringContent("{\"id\":\"durable-1\",\"version\":\"2\"}", Encoding.UTF8, "application/json") };
            }));
            byte[] payload = new byte[total]; GoogleDriveResumableUploadCheckpoint? saved = null;
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            using (var first = new GoogleDriveClient(Session(http, receipts))) {
                await Assert.ThrowsAsync<StopAfterCheckpointException>(() => first.UploadResumableStreamAsync(
                    new MemoryStream(payload, false), payload.Length,
                    new GoogleDriveUploadOptions { Name = "durable.bin", ResumableChunkSize = 256 * 1024 },
                    checkpointSink: (checkpoint, _) => {
                        saved = checkpoint;
                        if (checkpoint.ConfirmedBytes > 0) throw new StopAfterCheckpointException();
                        return Task.CompletedTask;
                    }));
            }
            Assert.DoesNotContain(receipts, receipt =>
                receipt.Succeeded && receipt.MutationKind == GoogleWorkspaceMutationKind.Create);
            GoogleDriveResumableUploadCheckpoint restored = GoogleDriveResumableUploadCheckpoint.Parse(Assert.IsType<GoogleDriveResumableUploadCheckpoint>(saved).Value);
            using var second = new GoogleDriveClient(Session(http, receipts));
            GoogleDriveResumableUploadResult result = await second.UploadResumableStreamAsync(
                new MemoryStream(payload, false), payload.Length,
                new GoogleDriveUploadOptions { Name = "durable.bin", ResumableChunkSize = 256 * 1024 }, restored);
            Assert.Equal("durable-1", result.File.Id);
            Assert.Equal(total, result.Checkpoint.ConfirmedBytes);
            Assert.DoesNotContain("upload.googleapis.com", result.Checkpoint.ToString(), StringComparison.Ordinal);
            Assert.DoesNotContain(receipts, receipt => receipt.Target.Contains("upload.googleapis.com", StringComparison.Ordinal));
            GoogleWorkspaceOperationReceipt createReceipt = Assert.Single(receipts, receipt =>
                receipt.Succeeded && receipt.MutationKind == GoogleWorkspaceMutationKind.Create);
            Assert.Equal(GoogleWorkspaceRevisionPreconditionKind.ResumableSessionState,
                createReceipt.RevisionPreconditionKind);
            Assert.StartsWith("resumable-session:", createReceipt.EnforcedRevision, StringComparison.Ordinal);
            Assert.Contains(receipts, receipt => receipt.Succeeded &&
                receipt.MutationKind == GoogleWorkspaceMutationKind.Action &&
                receipt.EnforcedRevision != null &&
                receipt.EnforcedRevision.StartsWith("content-range:bytes ", StringComparison.Ordinal));
        }

        [Fact]
        public async Task ResumableUploadReconcilesACommittedChunkWhenResponseBufferingFails() {
            const int total = 16;
            bool committed = false;
            int statusQueries = 0;
            using var http = new HttpClient(new Handler(request => {
                if (request.Method == HttpMethod.Post) {
                    var initiated = Json("{}");
                    initiated.Headers.Location = new Uri("https://upload.googleapis.com/ambiguous-read");
                    return initiated;
                }

                string range = request.Content!.Headers.GetValues("Content-Range").Single();
                if (range == $"bytes */{total}") {
                    statusQueries++;
                    return committed
                        ? new HttpResponseMessage(HttpStatusCode.OK) {
                            Content = new StringContent("{\"id\":\"reconciled-1\",\"version\":\"2\"}",
                                Encoding.UTF8, "application/json"),
                        }
                        : new HttpResponseMessage((HttpStatusCode)308) {
                            Content = new StringContent(string.Empty),
                        };
                }

                committed = true;
                return new HttpResponseMessage(HttpStatusCode.Created) {
                    Content = new ThrowingReadContent(),
                };
            }));
            using var client = new GoogleDriveClient(Session(http,
                new List<GoogleWorkspaceOperationReceipt>()));

            GoogleDriveFile file = await client.UploadResumableAsync(
                new byte[total], new GoogleDriveUploadOptions { Name = "ambiguous.bin" });

            Assert.Equal("reconciled-1", file.Id);
            Assert.Equal(1, statusQueries);
        }

        [Fact]
        public async Task ResumableUploadReconcilesACommittedChunkWhenElapsedBudgetExpiresAfterResponse() {
            const int total = 16;
            bool committed = false;
            int statusQueries = 0;
            using var http = new HttpClient(new Handler(async request => {
                if (request.Method == HttpMethod.Post) {
                    var initiated = Json("{}");
                    initiated.Headers.Location = new Uri("https://upload.googleapis.com/ambiguous-timeout");
                    return initiated;
                }

                string range = request.Content!.Headers.GetValues("Content-Range").Single();
                if (range == $"bytes */{total}") {
                    statusQueries++;
                    return new HttpResponseMessage(HttpStatusCode.OK) {
                        Content = new StringContent("{\"id\":\"reconciled-timeout\",\"version\":\"2\"}",
                            Encoding.UTF8, "application/json"),
                    };
                }

                committed = true;
                await Task.Delay(TimeSpan.FromMilliseconds(100)).ConfigureAwait(false);
                return new HttpResponseMessage(HttpStatusCode.Created) {
                    Content = new StringContent("{\"id\":\"already-created\",\"version\":\"2\"}",
                        Encoding.UTF8, "application/json"),
                };
            }));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            var sessionOptions = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                MaxRetryCount = 0,
                MaxRetryElapsedTime = TimeSpan.FromMilliseconds(50),
                OperationReceiptSink = receipts.Add,
            };
            sessionOptions.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                sessionOptions.ExpectedAccount!, context.RequiredScopes, context.Target,
                context.AdapterExpectedRevision ?? GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
                context.MaxRetryCount, context.MaxRetryElapsedTime, context.RateLimitPolicy,
                GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            using var client = new GoogleDriveClient(new GoogleWorkspaceSession(
                VerifiedCredentialSource(sessionOptions.ExpectedAccount!), sessionOptions));

            GoogleDriveFile file = await client.UploadResumableAsync(
                new byte[total], new GoogleDriveUploadOptions { Name = "ambiguous-timeout.bin" });

            Assert.True(committed);
            Assert.Equal("reconciled-timeout", file.Id);
            Assert.Equal(1, statusQueries);
        }

        [Fact]
        public async Task ResumableUploadDoesNotReconcileProgressCallbackFailures() {
            const int total = 16;
            int statusQueries = 0;
            using var http = new HttpClient(new Handler(request => {
                if (request.Method == HttpMethod.Post) {
                    var initiated = Json("{}");
                    initiated.Headers.Location = new Uri("https://upload.googleapis.com/progress-failure");
                    return initiated;
                }

                string range = request.Content!.Headers.GetValues("Content-Range").Single();
                if (range == $"bytes */{total}") statusQueries++;
                return new HttpResponseMessage(HttpStatusCode.Created) {
                    Content = new StringContent("{\"id\":\"created-before-progress-failure\",\"version\":\"2\"}",
                        Encoding.UTF8, "application/json"),
                };
            }));
            using var client = new GoogleDriveClient(Session(http,
                new List<GoogleWorkspaceOperationReceipt>()));

            await Assert.ThrowsAsync<IOException>(() => client.UploadResumableAsync(
                new byte[total], new GoogleDriveUploadOptions {
                    Name = "progress-failure.bin",
                    Progress = new ThrowingProgress(),
                }));

            Assert.Equal(0, statusQueries);
        }

        [Theory]
        [InlineData("not-a-range-15")]
        [InlineData("bytes=0-16")]
        [InlineData("bytes=1-15")]
        public async Task ResumableUploadRejectsMalformedOrOutOfBoundsConfirmedRange(string confirmedRange) {
            const int total = 16;
            using var http = new HttpClient(new Handler(request => {
                if (request.Method == HttpMethod.Post) {
                    var initiated = Json("{}");
                    initiated.Headers.Location = new Uri("https://upload.googleapis.com/invalid-range");
                    return initiated;
                }

                var status = new HttpResponseMessage((HttpStatusCode)308) {
                    Content = new StringContent(string.Empty),
                };
                status.Headers.TryAddWithoutValidation("Range", confirmedRange);
                return status;
            }));
            using var client = new GoogleDriveClient(Session(http,
                new List<GoogleWorkspaceOperationReceipt>()));

            await Assert.ThrowsAsync<InvalidDataException>(() =>
                client.UploadResumableStreamAsync(new MemoryStream(new byte[total]), total,
                    new GoogleDriveUploadOptions { Name = "invalid-range.bin" }));
        }

        [Fact]
        public async Task ResumableUploadRejectsSourceMutationBeforeReportingSuccess() {
            const int total = 16;
            using var http = new HttpClient(new Handler(request => {
                if (request.Method == HttpMethod.Post) {
                    var initiated = Json("{}");
                    initiated.Headers.Location = new Uri("https://upload.googleapis.com/mutable");
                    return initiated;
                }
                string range = request.Content!.Headers.GetValues("Content-Range").Single();
                return range == $"bytes */{total}"
                    ? new HttpResponseMessage((HttpStatusCode)308) { Content = new StringContent(string.Empty) }
                    : new HttpResponseMessage(HttpStatusCode.Created) { Content = new StringContent("{\"id\":\"mutable-1\",\"version\":\"2\"}", Encoding.UTF8, "application/json") };
            }));
            byte[] payload = new byte[total];
            using var client = new GoogleDriveClient(Session(http, new List<GoogleWorkspaceOperationReceipt>()));

            await Assert.ThrowsAsync<InvalidOperationException>(() => client.UploadResumableStreamAsync(
                new MemoryStream(payload), payload.Length, new GoogleDriveUploadOptions { Name = "mutable.bin" },
                checkpointSink: (checkpoint, _) => {
                    if (checkpoint.ConfirmedBytes == 0) payload[0] = 1;
                    return Task.CompletedTask;
                }));
        }

        [Fact]
        public async Task ResumableUploadDoesNotPublishCompletedCheckpointForChangedSource() {
            const int total = 16;
            byte[] payload = new byte[total];
            var confirmedOffsets = new List<long>();
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            using var http = new HttpClient(new Handler(async request => {
                if (request.Method == HttpMethod.Post) {
                    var initiated = Json("{}");
                    initiated.Headers.Location = new Uri("https://upload.googleapis.com/final-source-check");
                    return initiated;
                }
                string range = request.Content!.Headers.GetValues("Content-Range").Single();
                if (range == $"bytes */{total}") {
                    return new HttpResponseMessage((HttpStatusCode)308) {
                        Content = new StringContent(string.Empty),
                    };
                }
                _ = await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                payload[0] = 1;
                return new HttpResponseMessage(HttpStatusCode.Created) {
                    Content = new StringContent("{\"id\":\"changed-1\",\"version\":\"2\"}",
                        Encoding.UTF8, "application/json"),
                };
            }));
            using var client = new GoogleDriveClient(Session(http, receipts));

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                client.UploadResumableStreamAsync(new MemoryStream(payload), payload.Length,
                    new GoogleDriveUploadOptions { Name = "changed.bin" },
                    checkpointSink: (checkpoint, _) => {
                        confirmedOffsets.Add(checkpoint.ConfirmedBytes);
                        return Task.CompletedTask;
                    }));

            Assert.Equal(new long[] { 0 }, confirmedOffsets);
            GoogleWorkspaceOperationReceipt createReceipt = Assert.Single(receipts, receipt =>
                receipt.MutationKind == GoogleWorkspaceMutationKind.Create);
            Assert.True(createReceipt.Succeeded);
        }

        [Fact]
        public async Task ResumableUploadCheckpointRejectsChangedContentTypeBeforeResumeRequest() {
            int requests = 0;
            string? initiatedContentType = null;
            using var http = new HttpClient(new Handler(request => {
                requests++;
                initiatedContentType = request.Headers.TryGetValues("X-Upload-Content-Type", out var values)
                    ? values.Single()
                    : initiatedContentType;
                var response = Json("{}");
                response.Headers.Location = new Uri("https://upload.googleapis.com/content-type");
                return response;
            }));
            byte[] payload = Encoding.UTF8.GetBytes("durable payload");
            GoogleDriveResumableUploadCheckpoint? saved = null;
            var initialOptions = new GoogleDriveUploadOptions {
                Name = "payload.bin",
                ContentType = "application/octet-stream",
            };
            using var client = new GoogleDriveClient(Session(http,
                new List<GoogleWorkspaceOperationReceipt>(),
                policyObserver: _ => initialOptions.ContentType = "text/plain"));

            await Assert.ThrowsAsync<StopAfterCheckpointException>(() =>
                client.UploadResumableStreamAsync(new MemoryStream(payload, writable: false), payload.Length,
                    initialOptions,
                    checkpointSink: (checkpoint, _) => {
                        saved = checkpoint;
                        throw new StopAfterCheckpointException();
                    }));
            Assert.Equal(1, requests);
            Assert.Equal("application/octet-stream", initiatedContentType);

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                client.UploadResumableStreamAsync(new MemoryStream(payload, writable: false), payload.Length,
                    new GoogleDriveUploadOptions { Name = "payload.bin", ContentType = "text/plain" },
                    GoogleDriveResumableUploadCheckpoint.Parse(
                        Assert.IsType<GoogleDriveResumableUploadCheckpoint>(saved).Value)));
            Assert.Equal(1, requests);
        }

        [Fact]
        public async Task DurableDownloadTokenFailureDoesNotCreateDestinationOrParent() {
            using var http = new HttpClient(new Handler(_ =>
                Json("{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"3\"}")));
            string parent = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-token-" + Guid.NewGuid().ToString("N"));
            string path = Path.Combine(parent, "download.bin");
            try {
                var options = new GoogleWorkspaceSessionOptions { HttpClient = http };
                var session = new GoogleWorkspaceSession(new FailSecondCredentialSource(), options);
                using var client = new GoogleDriveClient(session);

                await Assert.ThrowsAsync<GoogleWorkspaceExportException>(() =>
                    client.DownloadToFileAsync("file-1", path));

                Assert.False(File.Exists(path));
                Assert.False(Directory.Exists(parent));
            } finally {
                if (Directory.Exists(parent)) Directory.Delete(parent, recursive: true);
            }
        }

        [Fact]
        public async Task DurableDownloadReturnsRecoverableStateWhenInitialCheckpointCannotBePersisted() {
            byte[] payload = Encoding.UTF8.GetBytes("abc");
            using var http = new HttpClient(new Handler(request => {
                if (!request.RequestUri!.Query.Contains("alt=media", StringComparison.Ordinal))
                    return Json("{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"3\"}");
                var response = new HttpResponseMessage(HttpStatusCode.PartialContent) {
                    Content = new ByteArrayContent(payload),
                };
                response.Content.Headers.ContentRange = new ContentRangeHeaderValue(0, 2, 3);
                return response;
            }));
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-uncheckpointed-" + Guid.NewGuid().ToString("N") + ".bin");
            try {
                using var client = new GoogleDriveClient(Session(http,
                    new List<GoogleWorkspaceOperationReceipt>()));

                GoogleDriveDownloadCheckpointPersistenceException exception =
                    await Assert.ThrowsAsync<GoogleDriveDownloadCheckpointPersistenceException>(() =>
                    client.DownloadToFileAsync("file-1", path,
                        checkpointSink: (_, _) => throw new StopAfterCheckpointException()));

                Assert.True(File.Exists(path));
                Assert.Equal(0, new FileInfo(path).Length);
                Assert.Equal(0, exception.Checkpoint.ConfirmedBytes);

                GoogleDriveDownloadCheckpoint completed = await client.DownloadToFileAsync(
                    "file-1", path, exception.Checkpoint);
                Assert.Equal(3, completed.ConfirmedBytes);
                Assert.Equal(payload, File.ReadAllBytes(path));
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task RangedDownloadCheckpointSurvivesRestartAndGuardsRevision() {
            byte[] payload = Enumerable.Range(0, 700_000).Select(value => unchecked((byte)value)).ToArray();
            using var http = new HttpClient(new Handler(request => {
                if (!request.RequestUri!.Query.Contains("alt=media", StringComparison.Ordinal)) return Json($"{{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"{payload.Length}\"}}");
                RangeItemHeaderValue range = request.Headers.Range!.Ranges.Single(); int start = checked((int)range.From!.Value); int end = checked((int)range.To!.Value);
                var response = new HttpResponseMessage(HttpStatusCode.PartialContent) { Content = new ByteArrayContent(payload.Skip(start).Take(end - start + 1).ToArray()) };
                response.Content.Headers.ContentRange = new ContentRangeHeaderValue(start, end, payload.Length);
                return response;
            }));
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-download-" + Guid.NewGuid().ToString("N") + ".bin");
            GoogleDriveDownloadCheckpoint? saved = null;
            try {
                using (var first = new GoogleDriveClient(Session(http, new List<GoogleWorkspaceOperationReceipt>()))) {
                    await Assert.ThrowsAsync<StopAfterCheckpointException>(() => first.DownloadToFileAsync("file-1", path,
                        checkpointSink: (checkpoint, _) => {
                            saved = checkpoint;
                            if (checkpoint.ConfirmedBytes > 0) throw new StopAfterCheckpointException();
                            return Task.CompletedTask;
                        },
                        chunkSize: 256 * 1024));
                }
                Assert.Equal(256 * 1024, new FileInfo(path).Length);
                GoogleDriveDownloadCheckpoint restored = GoogleDriveDownloadCheckpoint.Parse(Assert.IsType<GoogleDriveDownloadCheckpoint>(saved).Value);
                using var second = new GoogleDriveClient(Session(http, new List<GoogleWorkspaceOperationReceipt>()));
                GoogleDriveDownloadCheckpoint completed = await second.DownloadToFileAsync("file-1", path, restored, chunkSize: 256 * 1024);
                Assert.Equal(payload.Length, completed.ConfirmedBytes);
                Assert.Equal(payload, File.ReadAllBytes(path));
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        [Fact]
        public async Task RangedDownloadRejectsFinalDestinationContentMutationOnUnix() {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
            byte[] payload = Enumerable.Range(0, 300_000).Select(value => unchecked((byte)value)).ToArray();
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-final-mutation-" + Guid.NewGuid().ToString("N") + ".bin");
            int metadataRequests = 0;
            using var http = new HttpClient(new Handler(request => {
                if (!request.RequestUri!.Query.Contains("alt=media", StringComparison.Ordinal)) {
                    if (++metadataRequests == 2) {
                        using var tamper = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
                        tamper.WriteByte(unchecked((byte)(payload[0] + 1)));
                        tamper.Flush(flushToDisk: true);
                    }
                    return Json($"{{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"{payload.Length}\"}}");
                }
                RangeItemHeaderValue range = request.Headers.Range!.Ranges.Single();
                int start = checked((int)range.From!.Value);
                int end = checked((int)range.To!.Value);
                var response = new HttpResponseMessage(HttpStatusCode.PartialContent) {
                    Content = new ByteArrayContent(payload.Skip(start).Take(end - start + 1).ToArray()),
                };
                response.Content.Headers.ContentRange = new ContentRangeHeaderValue(start, end, payload.Length);
                return response;
            }));
            try {
                using var client = new GoogleDriveClient(Session(http,
                    new List<GoogleWorkspaceOperationReceipt>()));

                await Assert.ThrowsAsync<InvalidOperationException>(() =>
                    client.DownloadToFileAsync("file-1", path, chunkSize: 256 * 1024));

                Assert.NotEqual(payload, File.ReadAllBytes(path));
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        [Theory]
        [InlineData(HttpStatusCode.OK, false)]
        [InlineData(HttpStatusCode.PartialContent, true)]
        public async Task RangedDownloadRejectsUnconfirmedResponseBeforeWriting(
            HttpStatusCode statusCode, bool returnWrongRange) {
            const int total = 300_000;
            using var http = new HttpClient(new Handler(request => {
                if (!request.RequestUri!.Query.Contains("alt=media", StringComparison.Ordinal)) {
                    return Json($"{{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"{total}\"}}");
                }

                RangeItemHeaderValue requested = request.Headers.Range!.Ranges.Single();
                long start = requested.From!.Value;
                long end = requested.To!.Value;
                var response = new HttpResponseMessage(statusCode) {
                    Content = new ByteArrayContent(new byte[checked((int)(end - start + 1))]),
                };
                if (statusCode == HttpStatusCode.PartialContent) {
                    response.Content.Headers.ContentRange = returnWrongRange
                        ? new ContentRangeHeaderValue(start + 1, end + 1, total)
                        : new ContentRangeHeaderValue(start, end, total);
                }
                return response;
            }));
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-range-" + Guid.NewGuid().ToString("N") + ".bin");
            try {
                using var client = new GoogleDriveClient(Session(http,
                    new List<GoogleWorkspaceOperationReceipt>()));

                await Assert.ThrowsAsync<InvalidDataException>(() =>
                    client.DownloadToFileAsync("file-1", path, chunkSize: 256 * 1024));

                Assert.True(File.Exists(path));
                Assert.Equal(0, new FileInfo(path).Length);
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        [Fact]
        public async Task DurableDownloadRejectsMetadataAboveLimitBeforeCreatingDestination() {
            int mediaRequests = 0;
            using var http = new HttpClient(new Handler(request => {
                if (request.RequestUri!.Query.Contains("alt=media", StringComparison.Ordinal)) mediaRequests++;
                return Json("{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"5\"}");
            }));
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-limit-" + Guid.NewGuid().ToString("N") + ".bin");
            using var client = new GoogleDriveClient(
                Session(http, new List<GoogleWorkspaceOperationReceipt>()),
                new GoogleDriveClientOptions { MaxDownloadBytes = 4 });

            await Assert.ThrowsAsync<InvalidDataException>(() =>
                client.DownloadToFileAsync("file-1", path));

            Assert.False(File.Exists(path));
            Assert.Equal(0, mediaRequests);
        }

        [Fact]
        public async Task DurableDownloadRejectsNegativeMetadataSizeBeforeCreatingDestination() {
            int mediaRequests = 0;
            using var http = new HttpClient(new Handler(request => {
                if (request.RequestUri!.Query.Contains("alt=media", StringComparison.Ordinal)) mediaRequests++;
                return Json("{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"-1\"}");
            }));
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-negative-size-" + Guid.NewGuid().ToString("N") + ".bin");
            using var client = new GoogleDriveClient(
                Session(http, new List<GoogleWorkspaceOperationReceipt>()),
                new GoogleDriveClientOptions { MaxDownloadBytes = 4 });

            await Assert.ThrowsAsync<InvalidDataException>(() =>
                client.DownloadToFileAsync("file-1", path));

            Assert.False(File.Exists(path));
            Assert.Equal(0, mediaRequests);
        }

        [Fact]
        public async Task GuardedDownloadNeverTruncatesAnExistingDestination() {
            using var http = new HttpClient(new Handler(_ =>
                Json("{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"3\"}")));
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-existing-" + Guid.NewGuid().ToString("N") + ".bin");
            byte[] original = Encoding.UTF8.GetBytes("do not replace");
            File.WriteAllBytes(path, original);
            try {
                using var client = new GoogleDriveClient(Session(http, new List<GoogleWorkspaceOperationReceipt>()));

                await Assert.ThrowsAsync<IOException>(() => client.DownloadToFileAsync("file-1", path));

                Assert.Equal(original, File.ReadAllBytes(path));
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        [Fact]
        public async Task ZeroByteDownloadCheckpointRejectsAReplacementDestination() {
            using var http = new HttpClient(new Handler(_ =>
                Json("{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"3\"}")));
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-replaced-" + Guid.NewGuid().ToString("N") + ".bin");
            GoogleDriveDownloadCheckpoint? saved = null;
            byte[] replacement = Encoding.UTF8.GetBytes("unrelated content");
            try {
                using var stopAfterPersist = new CancellationTokenSource();
                using (var first = new GoogleDriveClient(Session(http,
                           new List<GoogleWorkspaceOperationReceipt>()))) {
                    await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                        first.DownloadToFileAsync("file-1", path,
                            checkpointSink: (checkpoint, _) => {
                                saved = checkpoint;
                                stopAfterPersist.Cancel();
                                return Task.CompletedTask;
                            },
                            chunkSize: 256 * 1024,
                            cancellationToken: stopAfterPersist.Token));
                }

                File.Delete(path);
                File.WriteAllBytes(path, replacement);
                GoogleDriveDownloadCheckpoint restored = GoogleDriveDownloadCheckpoint.Parse(
                    Assert.IsType<GoogleDriveDownloadCheckpoint>(saved).Value);
                using var second = new GoogleDriveClient(Session(http,
                    new List<GoogleWorkspaceOperationReceipt>()));

                await Assert.ThrowsAsync<InvalidOperationException>(() =>
                    second.DownloadToFileAsync("file-1", path, restored,
                        chunkSize: 256 * 1024));

                Assert.Equal(replacement, File.ReadAllBytes(path));
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

#if NET8_0_OR_GREATER
        [Fact]
        public async Task ZeroByteDownloadCheckpointRejectsHardLinkedDestination() {
            using var http = new HttpClient(new Handler(_ =>
                Json("{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"3\"}")));
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-linked-" + Guid.NewGuid().ToString("N") + ".bin");
            string alias = path + ".alias";
            GoogleDriveDownloadCheckpoint? saved = null;
            try {
                using var stopAfterPersist = new CancellationTokenSource();
                using (var first = new GoogleDriveClient(Session(http,
                           new List<GoogleWorkspaceOperationReceipt>()))) {
                    await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                        first.DownloadToFileAsync("file-1", path,
                            checkpointSink: (checkpoint, _) => {
                                saved = checkpoint;
                                stopAfterPersist.Cancel();
                                return Task.CompletedTask;
                            },
                            chunkSize: 256 * 1024,
                            cancellationToken: stopAfterPersist.Token));
                }

                CreateHardLinkForTest(path, alias);
                GoogleDriveDownloadCheckpoint restored = GoogleDriveDownloadCheckpoint.Parse(
                    Assert.IsType<GoogleDriveDownloadCheckpoint>(saved).Value);
                using var second = new GoogleDriveClient(Session(http,
                    new List<GoogleWorkspaceOperationReceipt>()));

                await Assert.ThrowsAsync<IOException>(() =>
                    second.DownloadToFileAsync("file-1", path, restored,
                        chunkSize: 256 * 1024));

                Assert.Equal(0, new FileInfo(path).Length);
                Assert.Equal(0, new FileInfo(alias).Length);
            } finally {
                if (File.Exists(alias)) File.Delete(alias);
                if (File.Exists(path)) File.Delete(path);
            }
        }

        private static void CreateHardLinkForTest(string existingPath, string linkPath) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                if (!CreateHardLinkWindows(linkPath, existingPath, IntPtr.Zero)) {
                    throw new IOException("The hard-link test fixture could not be created.");
                }
            } else if (CreateHardLinkUnix(existingPath, linkPath) != 0) {
                throw new IOException("The hard-link test fixture could not be created.");
            }
        }

        [DllImport("kernel32.dll", EntryPoint = "CreateHardLinkW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CreateHardLinkWindows(string fileName, string existingFileName,
            IntPtr securityAttributes);

        [DllImport("libc", EntryPoint = "link", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int CreateHardLinkUnix(string existingPath, string linkPath);
#endif

        private static GoogleWorkspaceSession Session(HttpClient http, IList<GoogleWorkspaceOperationReceipt> receipts,
            GoogleWorkspaceDataLossDecision dataLossDecision = GoogleWorkspaceDataLossDecision.RejectPotentialLoss,
            string? acceptedLoss = null,
            Action<GoogleWorkspaceOperationContext>? policyObserver = null) {
            var options = new GoogleWorkspaceSessionOptions { HttpClient = http, ExpectedAccount = "test@example.com", OperationReceiptSink = receipts.Add };
            options.OperationPolicyProvider = context => {
                policyObserver?.Invoke(context);
                return new GoogleWorkspaceOperationPolicy(options.ExpectedAccount!,
                    context.RequiredScopes, context.Target,
                    context.RevisionPreconditionKind switch {
                        GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate => GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
                        GoogleWorkspaceRevisionPreconditionKind.PayloadRevision => context.AdapterExpectedRevision!,
                        GoogleWorkspaceRevisionPreconditionKind.ResumableSessionState => context.AdapterExpectedRevision!,
                        GoogleWorkspaceRevisionPreconditionKind.Unavailable => GoogleWorkspaceOperationPolicy.ExplicitlyUnversionedRevision("test mutation"),
                        _ => "\"test-etag\"",
                    },
                    context.MaxRetryCount, context.MaxRetryElapsedTime, context.RateLimitPolicy,
                    dataLossDecision, acceptedLoss);
            };
            return new GoogleWorkspaceSession(VerifiedCredentialSource(options.ExpectedAccount!), options);
        }
        private static GoogleWorkspaceSessionOptions MutationOptions(HttpClient http,
            Action<GoogleWorkspaceOperationReceipt> receiptSink) {
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                OperationReceiptSink = receiptSink,
            };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, context.RequiredScopes, context.Target,
                GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
                context.MaxRetryCount, context.MaxRetryElapsedTime, context.RateLimitPolicy,
                GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            return options;
        }
        private static GoogleWorkspaceHttpTransport MutationTransport(
            GoogleWorkspaceSessionOptions options, params string[] scopes) {
            var session = new GoogleWorkspaceSession(VerifiedCredentialSource(options.ExpectedAccount!), options);
            _ = session.AcquireAccessTokenAsync(scopes).GetAwaiter().GetResult();
            return new GoogleWorkspaceHttpTransport(session);
        }
        private static IGoogleWorkspaceCredentialSource VerifiedCredentialSource(string account) =>
            new DelegateGoogleWorkspaceCredentialSource((scopes, _) => Task.FromResult(
                GoogleWorkspaceAccessToken.FromVerifiedCredential("token", DateTimeOffset.UtcNow.AddMinutes(5),
                    new GoogleWorkspaceCredentialBinding(account, scopes))));
        private static HttpResponseMessage Json(string value) => new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(value, Encoding.UTF8, "application/json") };
        private sealed class StopAfterCheckpointException : Exception { }
        private sealed class ThrowingProgress : IProgress<GoogleDriveTransferProgress> {
            public void Report(GoogleDriveTransferProgress value) =>
                throw new IOException("The progress observer failed.");
        }
        private sealed class FailSecondCredentialSource : IGoogleWorkspaceCredentialSource {
            private int _calls;
            public Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(
                IEnumerable<string> scopes, CancellationToken cancellationToken = default) {
                if (Interlocked.Increment(ref _calls) > 1) {
                    throw new InvalidOperationException("token acquisition failed");
                }
                return Task.FromResult(new GoogleWorkspaceAccessToken(
                    "token", DateTimeOffset.UtcNow.AddMinutes(5), scopes.ToArray()));
            }
        }
        private sealed class TrackingResponseMessage : HttpResponseMessage {
            internal TrackingResponseMessage(HttpStatusCode statusCode) : base(statusCode) { }
            internal bool WasDisposed { get; private set; }
            protected override void Dispose(bool disposing) {
                WasDisposed = true;
                base.Dispose(disposing);
            }
        }
        private sealed class ThrowingReadContent : HttpContent {
            protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) =>
                Task.FromException(new IOException("Transient response read failure."));
            protected override bool TryComputeLength(out long length) { length = 0; return false; }
            protected override Task<Stream> CreateContentReadStreamAsync() =>
                Task.FromResult<Stream>(new ThrowingReadStream());
        }
        private sealed class ThrowingReadStream : MemoryStream {
            public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
                CancellationToken cancellationToken) =>
                Task.FromException<int>(new IOException("Transient response read failure."));
        }
        private sealed class Handler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;
            internal Handler(Func<HttpRequestMessage, HttpResponseMessage> handler) { _handler = request => Task.FromResult(handler(request)); }
            internal Handler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) { _handler = handler; }
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => _handler(request);
        }
    }
}
