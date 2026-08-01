using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization.Metadata;

namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Dependency-light HTTP transport shared by Google Workspace domain packages.
    /// </summary>
    public sealed class GoogleWorkspaceHttpTransport : IDisposable {
        private const long MaximumErrorResponseBytes = 64L * 1024;
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions {
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = null,
            WriteIndented = false,
        };

        private readonly GoogleWorkspaceSessionOptions _options;
        private readonly HttpClient _client;
        private readonly bool _ownsClient;
        private bool _disposed;

        public GoogleWorkspaceHttpTransport(GoogleWorkspaceSessionOptions options) {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _ownsClient = options.HttpClient == null;
            _client = options.HttpClient ?? new HttpClient();
            if (_ownsClient) {
                _client.Timeout = System.Threading.Timeout.InfiniteTimeSpan;
            }
        }

        [RequiresUnreferencedCode("Use the overload that accepts JsonTypeInfo<TResponse> in trimmed applications.")]
        [RequiresDynamicCode("Use the overload that accepts JsonTypeInfo<TResponse> in NativeAOT applications.")]
        public Task<TResponse> SendJsonAsync<TResponse>(
            string accessToken,
            HttpMethod method,
            string uri,
            object? payload,
            GoogleWorkspaceRequestSafety requestSafety,
            string serviceName,
            TranslationReport report,
            CancellationToken cancellationToken = default,
            long? maxResponseBytes = null,
            GoogleWorkspaceMutationKind mutationKind = GoogleWorkspaceMutationKind.Unspecified,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition = null,
            IReadOnlyCollection<string>? requiredScopes = null,
            bool potentialDataLoss = false) {
            return SendAsync<TResponse>(
                accessToken,
                method,
                uri,
                payload == null
                    ? null
                    : (() => new StringContent(JsonSerializer.Serialize(payload, JsonOptions), Encoding.UTF8, "application/json")),
                requestSafety,
                serviceName,
                report,
                cancellationToken,
                maxResponseBytes,
                mutationKind,
                revisionPrecondition,
                requiredScopes,
                potentialDataLoss);
        }

        /// <summary>
        /// Sends a typed JSON payload and deserializes the response with source-generated metadata.
        /// </summary>
        public Task<TResponse> SendJsonAsync<TRequest, TResponse>(
            string accessToken,
            HttpMethod method,
            string uri,
            TRequest payload,
            GoogleWorkspaceRequestSafety requestSafety,
            string serviceName,
            TranslationReport report,
            JsonTypeInfo<TRequest> requestTypeInfo,
            JsonTypeInfo<TResponse> responseTypeInfo,
            CancellationToken cancellationToken = default,
            long? maxResponseBytes = null,
            GoogleWorkspaceMutationKind mutationKind = GoogleWorkspaceMutationKind.Unspecified,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition = null,
            IReadOnlyCollection<string>? requiredScopes = null,
            bool potentialDataLoss = false) {
            if (requestTypeInfo == null) throw new ArgumentNullException(nameof(requestTypeInfo));
            if (responseTypeInfo == null) throw new ArgumentNullException(nameof(responseTypeInfo));
            return SendAsync(
                accessToken,
                method,
                uri,
                () => new StringContent(JsonSerializer.Serialize(payload, requestTypeInfo), Encoding.UTF8, "application/json"),
                requestSafety,
                serviceName,
                report,
                responseTypeInfo,
                cancellationToken,
                maxResponseBytes,
                mutationKind,
                revisionPrecondition,
                requiredScopes,
                potentialDataLoss);
        }

        /// <summary>
        /// Sends an optional JSON node and deserializes the response with source-generated metadata.
        /// </summary>
        public Task<TResponse> SendJsonAsync<TResponse>(
            string accessToken,
            HttpMethod method,
            string uri,
            JsonNode? payload,
            GoogleWorkspaceRequestSafety requestSafety,
            string serviceName,
            TranslationReport report,
            JsonTypeInfo<TResponse> responseTypeInfo,
            CancellationToken cancellationToken = default,
            long? maxResponseBytes = null,
            GoogleWorkspaceMutationKind mutationKind = GoogleWorkspaceMutationKind.Unspecified,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition = null,
            IReadOnlyCollection<string>? requiredScopes = null,
            bool potentialDataLoss = false) {
            if (responseTypeInfo == null) throw new ArgumentNullException(nameof(responseTypeInfo));
            return SendAsync(
                accessToken,
                method,
                uri,
                payload == null
                    ? null
                    : (() => new StringContent(payload.ToJsonString(JsonOptions), Encoding.UTF8, "application/json")),
                requestSafety,
                serviceName,
                report,
                responseTypeInfo,
                cancellationToken,
                maxResponseBytes,
                mutationKind,
                revisionPrecondition,
                requiredScopes,
                potentialDataLoss);
        }

        [RequiresUnreferencedCode("Use the overload that accepts JsonTypeInfo<TResponse> in trimmed applications.")]
        [RequiresDynamicCode("Use the overload that accepts JsonTypeInfo<TResponse> in NativeAOT applications.")]
        public Task<TResponse> SendAsync<TResponse>(
            string accessToken,
            HttpMethod method,
            string uri,
            Func<HttpContent?>? contentFactory,
            GoogleWorkspaceRequestSafety requestSafety,
            string serviceName,
            TranslationReport report,
            CancellationToken cancellationToken = default,
            long? maxResponseBytes = null,
            GoogleWorkspaceMutationKind mutationKind = GoogleWorkspaceMutationKind.Unspecified,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition = null,
            IReadOnlyCollection<string>? requiredScopes = null,
            bool potentialDataLoss = false) {
            return SendAsyncCore(
                accessToken,
                method,
                uri,
                contentFactory,
                requestSafety,
                serviceName,
                report,
                body => JsonSerializer.Deserialize<TResponse>(body, JsonOptions),
                cancellationToken,
                maxResponseBytes,
                mutationKind,
                revisionPrecondition,
                requiredScopes,
                potentialDataLoss);
        }

        /// <summary>
        /// Sends a request and deserializes the response with source-generated metadata.
        /// </summary>
        public Task<TResponse> SendAsync<TResponse>(
            string accessToken,
            HttpMethod method,
            string uri,
            Func<HttpContent?>? contentFactory,
            GoogleWorkspaceRequestSafety requestSafety,
            string serviceName,
            TranslationReport report,
            JsonTypeInfo<TResponse> responseTypeInfo,
            CancellationToken cancellationToken = default,
            long? maxResponseBytes = null,
            GoogleWorkspaceMutationKind mutationKind = GoogleWorkspaceMutationKind.Unspecified,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition = null,
            IReadOnlyCollection<string>? requiredScopes = null,
            bool potentialDataLoss = false) {
            if (responseTypeInfo == null) throw new ArgumentNullException(nameof(responseTypeInfo));
            return SendAsyncCore(
                accessToken,
                method,
                uri,
                contentFactory,
                requestSafety,
                serviceName,
                report,
                body => JsonSerializer.Deserialize(body, responseTypeInfo),
                cancellationToken,
                maxResponseBytes,
                mutationKind,
                revisionPrecondition,
                requiredScopes,
                potentialDataLoss);
        }

        private async Task<TResponse> SendAsyncCore<TResponse>(
            string accessToken,
            HttpMethod method,
            string uri,
            Func<HttpContent?>? contentFactory,
            GoogleWorkspaceRequestSafety requestSafety,
            string serviceName,
            TranslationReport report,
            Func<string, TResponse?> deserialize,
            CancellationToken cancellationToken,
            long? maxResponseBytes,
            GoogleWorkspaceMutationKind mutationKind,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition,
            IReadOnlyCollection<string>? requiredScopes,
            bool potentialDataLoss) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(accessToken)) throw new ArgumentException("Access token is required.", nameof(accessToken));
            if (method == null) throw new ArgumentNullException(nameof(method));
            if (string.IsNullOrWhiteSpace(uri)) throw new ArgumentException("Request URI is required.", nameof(uri));
            if (string.IsNullOrWhiteSpace(serviceName)) throw new ArgumentException("Service name is required.", nameof(serviceName));
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (maxResponseBytes.HasValue && maxResponseBytes.Value <= 0) throw new ArgumentOutOfRangeException(nameof(maxResponseBytes));

            string effectiveUri = AppendQueryParameter(uri, "quotaUser", _options.QuotaUser);
            string visibleTarget = SanitizeDiagnosticTarget(uri);
            string? requestId = _options.RequestIdFactory?.Invoke();
            var retryOptions = GoogleWorkspaceRetryOptions.FromSessionOptions(_options);
            TimeSpan requestTimeout = _options.RequestTimeout;
            MutationAttempt? mutation = BeginMutation(method, visibleTarget, requestSafety, mutationKind,
                revisionPrecondition, serviceName, requestId, requiredScopes, retryOptions, potentialDataLoss);

            try {
                TResponse result = await GoogleWorkspaceRetryPolicy.SendAndProcessAsync(
                    _client,
                    () => {
                        HttpRequestMessage request = CreateRequest(accessToken, method, effectiveUri, contentFactory, requestId);
                        mutation?.ApplyRevisionPrecondition(request);
                        return request;
                    },
                    retryOptions,
                    requestSafety,
                    requestTimeout,
                    cancellationToken,
                    async (response, responseToken) => {
                    if (!response.IsSuccessStatusCode) {
                        byte[] errorBytes = await ReadResponseBytesAsync(
                            response.Content,
                            MaximumErrorResponseBytes,
                            responseToken,
                            truncateAtLimit: true).ConfigureAwait(false);
                        string errorBody = Encoding.UTF8.GetString(errorBytes);
                        throw GoogleWorkspaceApiException.Create(serviceName, method, visibleTarget,
                            response.StatusCode, errorBody);
                    }

                    byte[] responseBytes = await ReadResponseBytesAsync(
                        response.Content,
                        maxResponseBytes,
                        responseToken).ConfigureAwait(false);
                    string body = Encoding.UTF8.GetString(responseBytes);
                    if (typeof(TResponse) == typeof(object) || string.IsNullOrWhiteSpace(body)) {
                        return default!;
                    }

                    var result = deserialize(body);
                    if (result == null) {
                        throw new InvalidOperationException(
                            $"{serviceName} response from '{visibleTarget}' could not be deserialized.");
                    }

                    return result;
                    },
                    retryEvent => { mutation?.CountRetry(); ReportRetry(report, serviceName, retryEvent, visibleTarget); })
                    .ConfigureAwait(false);
                CompleteMutationSuccess(mutation);
                return result;
            } catch (Exception exception) {
                CompleteMutationFailure(mutation, exception);
                throw;
            }
        }

        public async Task<byte[]> SendBytesAsync(
            string accessToken,
            HttpMethod method,
            string uri,
            GoogleWorkspaceRequestSafety requestSafety,
            string serviceName,
            TranslationReport report,
            CancellationToken cancellationToken = default,
            bool preserveRequestUri = false,
            bool includeAuthorization = true,
            long? maxResponseBytes = null,
            Action<HttpRequestMessage>? configureRequest = null,
            GoogleWorkspaceMutationKind mutationKind = GoogleWorkspaceMutationKind.Unspecified,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition = null,
            Action<HttpResponseMessage>? validateResponse = null,
            IReadOnlyCollection<string>? requiredScopes = null,
            bool potentialDataLoss = false) {
            ThrowIfDisposed();
            if (maxResponseBytes.HasValue && maxResponseBytes.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(maxResponseBytes));
            }
            string effectiveUri = preserveRequestUri
                ? uri
                : AppendQueryParameter(uri, "quotaUser", _options.QuotaUser);
            string visibleTarget = SanitizeDiagnosticTarget(uri);
            string? requestId = _options.RequestIdFactory?.Invoke();
            var retryOptions = GoogleWorkspaceRetryOptions.FromSessionOptions(_options);
            TimeSpan requestTimeout = _options.RequestTimeout;
            MutationAttempt? mutation = BeginMutation(method, visibleTarget, requestSafety, mutationKind,
                revisionPrecondition, serviceName, requestId, requiredScopes, retryOptions, potentialDataLoss);

            try {
                byte[] result = await GoogleWorkspaceRetryPolicy.SendAndProcessAsync(
                    _client,
                    () => {
                        HttpRequestMessage request = CreateRequest(accessToken, method, effectiveUri, null, requestId, includeAuthorization);
                        configureRequest?.Invoke(request);
                        mutation?.ApplyRevisionPrecondition(request);
                        return request;
                    },
                    retryOptions,
                    requestSafety,
                    requestTimeout,
                    cancellationToken,
                    async (response, responseToken) => {
                    if (!response.IsSuccessStatusCode) {
                        byte[] errorBytes = await ReadResponseBytesAsync(
                            response.Content,
                            MaximumErrorResponseBytes,
                            responseToken,
                            truncateAtLimit: true).ConfigureAwait(false);
                        string body = Encoding.UTF8.GetString(errorBytes);
                        throw GoogleWorkspaceApiException.Create(serviceName,
                            method, visibleTarget, response.StatusCode, body);
                    }

                    validateResponse?.Invoke(response);
                    return await ReadResponseBytesAsync(response.Content,
                        maxResponseBytes, responseToken).ConfigureAwait(false);
                    },
                    retryEvent => { mutation?.CountRetry(); ReportRetry(report, serviceName, retryEvent, visibleTarget); })
                    .ConfigureAwait(false);
                CompleteMutationSuccess(mutation);
                return result;
            } catch (Exception exception) {
                CompleteMutationFailure(mutation, exception);
                throw;
            }
        }

        private static async Task<byte[]> ReadResponseBytesAsync(
            HttpContent content,
            long? maxResponseBytes,
            CancellationToken cancellationToken,
            bool truncateAtLimit = false) {
            long? limit = maxResponseBytes;
            if (limit.HasValue && content.Headers.ContentLength is long declaredLength
                && declaredLength > limit.Value && !truncateAtLimit) {
                throw new InvalidDataException(
                    $"The response declared {declaredLength} bytes, exceeding the configured limit of {limit.Value} bytes.");
            }

            using Stream input = await content.ReadAsStreamAsync().ConfigureAwait(false);
            using var output = new MemoryStream();
            byte[] buffer = new byte[81920];
            long total = 0;
            while (true) {
                int read = await input.ReadAsync(buffer, 0, buffer.Length,
                    cancellationToken).ConfigureAwait(false);
                if (read == 0) break;
                if (limit.HasValue && read > limit.Value - total) {
                    if (truncateAtLimit) {
                        output.Write(buffer, 0, checked((int)(limit.Value - total)));
                        break;
                    }
                    throw new InvalidDataException(
                        $"The response exceeded the configured limit of {limit.Value} bytes.");
                }
                output.Write(buffer, 0, read);
                total += read;
                if (truncateAtLimit && limit.HasValue && total == limit.Value) break;
            }
            return output.ToArray();
        }

        public async Task<GoogleWorkspaceHttpResponse> SendRawAsync(
            string accessToken,
            HttpMethod method,
            string uri,
            Func<HttpContent?>? contentFactory,
            GoogleWorkspaceRequestSafety requestSafety,
            string serviceName,
            TranslationReport report,
            CancellationToken cancellationToken = default,
            Action<HttpRequestMessage>? configureRequest = null,
            IReadOnlyCollection<HttpStatusCode>? additionalSuccessStatusCodes = null,
            bool preserveRequestUri = false,
            string? diagnosticTarget = null,
            GoogleWorkspaceMutationKind mutationKind = GoogleWorkspaceMutationKind.Unspecified,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition = null,
            IReadOnlyCollection<string>? requiredScopes = null,
            bool potentialDataLoss = false) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(accessToken)) throw new ArgumentException("Access token is required.", nameof(accessToken));
            if (method == null) throw new ArgumentNullException(nameof(method));
            if (string.IsNullOrWhiteSpace(uri)) throw new ArgumentException("Request URI is required.", nameof(uri));
            if (string.IsNullOrWhiteSpace(serviceName)) throw new ArgumentException("Service name is required.", nameof(serviceName));
            if (report == null) throw new ArgumentNullException(nameof(report));

            string effectiveUri = preserveRequestUri
                ? uri
                : AppendQueryParameter(uri, "quotaUser", _options.QuotaUser);
            string visibleTarget = string.IsNullOrWhiteSpace(diagnosticTarget)
                ? SanitizeDiagnosticTarget(uri)
                : diagnosticTarget!;
            string? requestId = _options.RequestIdFactory?.Invoke();
            var retryOptions = GoogleWorkspaceRetryOptions.FromSessionOptions(_options);
            TimeSpan requestTimeout = _options.RequestTimeout;
            MutationAttempt? mutation = BeginMutation(method, visibleTarget, requestSafety, mutationKind,
                revisionPrecondition, serviceName, requestId, requiredScopes, retryOptions, potentialDataLoss);

            try {
                using (var response = await GoogleWorkspaceRetryPolicy.SendAsync(
                    _client,
                    () => {
                        var request = CreateRequest(accessToken, method, effectiveUri, contentFactory, requestId);
                        configureRequest?.Invoke(request);
                        mutation?.ApplyRevisionPrecondition(request);
                        return request;
                    },
                    retryOptions,
                    requestSafety,
                    requestTimeout,
                    cancellationToken,
                    retryEvent => { mutation?.CountRetry(); ReportRetry(report, serviceName, retryEvent, visibleTarget); }).ConfigureAwait(false)) {
                    byte[] body = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                    bool accepted = response.IsSuccessStatusCode
                        || (additionalSuccessStatusCodes != null && additionalSuccessStatusCodes.Contains(response.StatusCode));
                    if (!accepted) {
                        string responseText = Encoding.UTF8.GetString(body);
                        throw GoogleWorkspaceApiException.Create(serviceName, method, visibleTarget, response.StatusCode, responseText);
                    }

                    var headers = response.Headers
                        .Concat(response.Content.Headers)
                        .GroupBy(header => header.Key, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(group => group.Key,
                            group => (IReadOnlyList<string>)group.SelectMany(header => header.Value).ToArray(),
                            StringComparer.OrdinalIgnoreCase);
                    var result = new GoogleWorkspaceHttpResponse(response.StatusCode, body,
                        response.Content.Headers.ContentType?.MediaType, headers);
                    CompleteMutationSuccess(mutation);
                    return result;
                }
            } catch (Exception exception) {
                CompleteMutationFailure(mutation, exception);
                throw;
            }
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }

            if (_ownsClient) {
                _client.Dispose();
            }

            _disposed = true;
        }

        internal DeferredMutation BeginDeferredMutation(HttpMethod method, string target,
            GoogleWorkspaceRequestSafety requestSafety, GoogleWorkspaceMutationKind mutationKind,
            GoogleWorkspaceRevisionPrecondition revisionPrecondition, string serviceName,
            IReadOnlyCollection<string> requiredScopes, bool potentialDataLoss = false) {
            ThrowIfDisposed();
            if (method == null) throw new ArgumentNullException(nameof(method));
            if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("The mutation target is required.", nameof(target));
            if (revisionPrecondition == null) throw new ArgumentNullException(nameof(revisionPrecondition));
            if (string.IsNullOrWhiteSpace(serviceName)) throw new ArgumentException("Service name is required.", nameof(serviceName));
            string? requestId = _options.RequestIdFactory?.Invoke();
            var retryOptions = GoogleWorkspaceRetryOptions.FromSessionOptions(_options);
            MutationAttempt mutation = BeginMutation(method, target, requestSafety, mutationKind,
                revisionPrecondition, serviceName, requestId, requiredScopes, retryOptions, potentialDataLoss)
                ?? throw new InvalidOperationException("A deferred mutation must not use safe request semantics.");
            return new DeferredMutation(mutation);
        }

        private HttpRequestMessage CreateRequest(
            string accessToken,
            HttpMethod method,
            string uri,
            Func<HttpContent?>? contentFactory,
            string? requestId,
            bool includeAuthorization = true) {
            var request = new HttpRequestMessage(method, uri);
            if (includeAuthorization) {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            }
            request.Headers.UserAgent.ParseAdd(BuildUserAgent(_options.ApplicationName));
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            if (!string.IsNullOrWhiteSpace(_options.QuotaProject)) {
                request.Headers.TryAddWithoutValidation("X-Goog-User-Project", _options.QuotaProject);
            }

            if (!string.IsNullOrWhiteSpace(requestId)) {
                request.Headers.TryAddWithoutValidation("X-Request-Id", requestId);
            }

            request.Content = contentFactory?.Invoke();
            return request;
        }

        private void ReportRetry(TranslationReport report, string serviceName, GoogleWorkspaceRetryEvent retryEvent,
            string? visibleTarget = null) {
            string target = visibleTarget ?? retryEvent.Uri;
            GoogleWorkspaceDiagnosticsDispatcher.AddUnique(
                report,
                _options,
                TranslationSeverity.Info,
                "ApiRetries",
                $"{serviceName} retried {retryEvent.Method} {target} after transient {retryEvent.Trigger} using {retryEvent.DelayStrategy} ({retryEvent.Delay.TotalMilliseconds:0} ms, retry {retryEvent.RetryAttempt} of {retryEvent.MaxRetryCount}).",
                $"{retryEvent.Method} {target}",
                code: GoogleWorkspaceDiagnosticCodes.ApiRetry);
        }

        private MutationAttempt? BeginMutation(HttpMethod method, string target,
            GoogleWorkspaceRequestSafety requestSafety, GoogleWorkspaceMutationKind mutationKind,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition, string service, string? requestId,
            IReadOnlyCollection<string>? requiredScopes, GoogleWorkspaceRetryOptions retryOptions,
            bool adapterDeclaredPotentialDataLoss) {
            if (requestSafety == GoogleWorkspaceRequestSafety.Safe) return null;
            string? expectedAccount = _options.ExpectedAccount;
            Func<GoogleWorkspaceOperationContext, GoogleWorkspaceOperationPolicy>? policyProvider =
                _options.OperationPolicyProvider;
            Action<GoogleWorkspaceOperationReceipt>? receiptSink = _options.OperationReceiptSink;
            string[] scopeSnapshot = requiredScopes?
                .Where(scope => !string.IsNullOrWhiteSpace(scope))
                .Distinct(StringComparer.Ordinal)
                .ToArray() ?? Array.Empty<string>();
            if (scopeSnapshot.Length == 0) {
                throw new InvalidOperationException("Google mutations require the adapter to declare the OAuth scopes requested for the operation.");
            }
            if (string.IsNullOrWhiteSpace(expectedAccount)) {
                throw new InvalidOperationException("Google mutations require GoogleWorkspaceSessionOptions.ExpectedAccount.");
            }
            if (policyProvider == null) {
                throw new InvalidOperationException("Google mutations require an explicit OperationPolicyProvider.");
            }
            if (receiptSink == null) {
                throw new InvalidOperationException("Google mutations require an OperationReceiptSink so outcomes are recorded.");
            }
            if (mutationKind == GoogleWorkspaceMutationKind.Unspecified) {
                if (method == HttpMethod.Delete) mutationKind = GoogleWorkspaceMutationKind.Delete;
                else if (method == HttpMethod.Put || method.Method == "PATCH") mutationKind = GoogleWorkspaceMutationKind.Update;
                else throw new InvalidOperationException("Google mutations whose semantics are not implied by PUT, PATCH, or DELETE require an explicit GoogleWorkspaceMutationKind.");
            }
            if (revisionPrecondition == null) {
                if (mutationKind == GoogleWorkspaceMutationKind.Create) {
                    revisionPrecondition = GoogleWorkspaceRevisionPrecondition.ResourceAbsentCreate;
                } else if (method == HttpMethod.Put || method.Method == "PATCH" || method == HttpMethod.Delete) {
                    revisionPrecondition = GoogleWorkspaceRevisionPrecondition.HttpEntityTag;
                } else {
                    throw new InvalidOperationException("Google mutations whose revision is not enforced by PUT, PATCH, or DELETE If-Match require an explicit GoogleWorkspaceRevisionPrecondition.");
                }
            }
            if (mutationKind == GoogleWorkspaceMutationKind.Create
                && revisionPrecondition.Kind != GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate
                && revisionPrecondition.Kind != GoogleWorkspaceRevisionPreconditionKind.ResumableSessionState) {
                throw new InvalidOperationException("An adapter-declared create must use a resource-absent or resumable-session revision precondition.");
            }
            bool potentialDataLoss = mutationKind == GoogleWorkspaceMutationKind.Delete
                || adapterDeclaredPotentialDataLoss;
            var context = new GoogleWorkspaceOperationContext(service, method.Method, target,
                requestSafety, mutationKind, revisionPrecondition, potentialDataLoss, requestId,
                Array.AsReadOnly(scopeSnapshot), retryOptions.MaxRetryCount, retryOptions.MaxElapsedTime,
                retryOptions.RateLimitPolicy);
            GoogleWorkspaceOperationPolicy policy = policyProvider(context)
                ?? throw new InvalidOperationException("The Google operation policy provider returned no policy.");
            if (!StringComparer.OrdinalIgnoreCase.Equals(policy.Account, expectedAccount)) {
                throw new InvalidOperationException("The Google operation policy account does not match the configured session account.");
            }
            if (policy.Scopes.Count != scopeSnapshot.Length
                || !new HashSet<string>(policy.Scopes, StringComparer.Ordinal).SetEquals(scopeSnapshot)) {
                throw new InvalidOperationException("The Google operation policy scopes do not match the OAuth scopes requested by the adapter.");
            }
            if (!StringComparer.Ordinal.Equals(policy.Target, target)) {
                throw new InvalidOperationException("The Google operation policy target does not match the request target.");
            }
            if (policy.MaxRetryCount != retryOptions.MaxRetryCount ||
                policy.MaxRetryElapsedTime != retryOptions.MaxElapsedTime ||
                policy.RateLimitPolicy != retryOptions.RateLimitPolicy) {
                throw new InvalidOperationException("The Google operation policy retry or rate-limit decision does not match the configured session behavior.");
            }
            if (context.PotentialDataLoss &&
                policy.DataLossDecision != GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss) {
                throw new InvalidOperationException("The Google operation policy rejects the potential data loss of this mutation.");
            }
            switch (revisionPrecondition.Kind) {
                case GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate:
                    if (context.MutationKind != GoogleWorkspaceMutationKind.Create
                        || !StringComparer.Ordinal.Equals(policy.ExpectedRevision,
                            GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision)) {
                        throw new InvalidOperationException("An adapter-declared create requires the resource-absent revision decision.");
                    }
                    break;
                case GoogleWorkspaceRevisionPreconditionKind.Unavailable:
                    if (!GoogleWorkspaceOperationPolicy.IsExplicitlyUnversioned(policy.ExpectedRevision)
                        || policy.DataLossDecision != GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss) {
                        throw new InvalidOperationException("A Google mutation with no usable revision precondition requires an explicitly unversioned revision and an accepted, named data-loss decision.");
                    }
                    break;
                case GoogleWorkspaceRevisionPreconditionKind.PayloadRevision:
                case GoogleWorkspaceRevisionPreconditionKind.ResumableSessionState:
                    if (!StringComparer.Ordinal.Equals(policy.ExpectedRevision,
                            revisionPrecondition.AdapterExpectedRevision)) {
                        throw new InvalidOperationException("The Google operation policy expected revision does not match the revision or resumable-session state enforced by the adapter request.");
                    }
                    break;
                case GoogleWorkspaceRevisionPreconditionKind.HttpEntityTag:
                    if (!EntityTagHeaderValue.TryParse(policy.ExpectedRevision, out EntityTagHeaderValue? entityTag)
                        || StringComparer.Ordinal.Equals(entityTag.Tag, "*") || entityTag.IsWeak) {
                        throw new InvalidOperationException("An HTTP revision precondition requires a strong, non-wildcard entity tag.");
                    }
                    break;
                default:
                    throw new InvalidOperationException("The Google mutation revision precondition is unspecified.");
            }
            return new MutationAttempt(policy, context, receiptSink);
        }

        private static void CompleteMutationSuccess(MutationAttempt? mutation) {
            GoogleWorkspaceReceiptPersistenceException? persistenceFailure = mutation?.Complete(true, "completed");
            if (persistenceFailure != null) {
                throw persistenceFailure;
            }
        }

        private static void CompleteMutationFailure(MutationAttempt? mutation, Exception operationFailure) {
            GoogleWorkspaceReceiptPersistenceException? persistenceFailure =
                mutation?.Complete(false, operationFailure.GetType().Name);
            if (persistenceFailure != null) {
                operationFailure.Data[GoogleWorkspaceReceiptPersistenceException.ExceptionDataKey] = persistenceFailure;
            }
        }

        internal sealed class MutationAttempt {
            private readonly GoogleWorkspaceOperationPolicy _policy;
            private readonly GoogleWorkspaceOperationContext _context;
            private readonly Action<GoogleWorkspaceOperationReceipt> _sink;
            private int _retryCount;
            private bool _completed;
            internal MutationAttempt(GoogleWorkspaceOperationPolicy policy, GoogleWorkspaceOperationContext context,
                Action<GoogleWorkspaceOperationReceipt> sink) { _policy = policy; _context = context; _sink = sink; }
            internal void CountRetry() { _retryCount++; }
            internal void ApplyRevisionPrecondition(HttpRequestMessage request) {
                if (_context.RevisionPreconditionKind != GoogleWorkspaceRevisionPreconditionKind.HttpEntityTag) {
                    return;
                }
                request.Headers.IfMatch.Clear();
                request.Headers.IfMatch.Add(EntityTagHeaderValue.Parse(_policy.ExpectedRevision));
            }
            internal GoogleWorkspaceReceiptPersistenceException? Complete(bool succeeded, string outcome) {
                if (_completed) return null;
                _completed = true;
                string? enforcedRevision = _context.RevisionPreconditionKind == GoogleWorkspaceRevisionPreconditionKind.HttpEntityTag
                    ? _policy.ExpectedRevision
                    : _context.AdapterExpectedRevision;
                var receipt = new GoogleWorkspaceOperationReceipt(_policy, _context.Service, _context.Method,
                    _context.Target, _context.RequestId, _retryCount, succeeded, outcome,
                    _context.MutationKind, _context.RevisionPreconditionKind, enforcedRevision);
                try {
                    _sink(receipt);
                    return null;
                } catch (Exception exception) {
                    return new GoogleWorkspaceReceiptPersistenceException(receipt, succeeded, exception);
                }
            }
        }

        internal sealed class DeferredMutation {
            private readonly MutationAttempt _mutation;

            internal DeferredMutation(MutationAttempt mutation) {
                _mutation = mutation;
            }

            internal void CompleteSuccess() => CompleteMutationSuccess(_mutation);

            internal void CompleteFailure(Exception exception) =>
                CompleteMutationFailure(_mutation, exception);
        }

        private static string SanitizeDiagnosticTarget(string uri) {
            int query = uri.IndexOf('?');
            int fragment = uri.IndexOf('#');
            int separator = query < 0 ? fragment : fragment < 0 ? query : Math.Min(query, fragment);
            return separator < 0 ? uri : uri.Substring(0, separator);
        }

        private static string AppendQueryParameter(string uri, string name, string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return uri;
            }

            string separator = uri.IndexOf('?') >= 0 ? "&" : "?";
            return uri + separator + Uri.EscapeDataString(name) + "=" + Uri.EscapeDataString(value!);
        }

        private static string BuildUserAgent(string applicationName) {
            var builder = new StringBuilder();
            foreach (char character in string.IsNullOrWhiteSpace(applicationName) ? "OfficeIMO" : applicationName) {
                if (char.IsLetterOrDigit(character) || character == '-' || character == '_' || character == '.') {
                    builder.Append(character);
                } else if (builder.Length == 0 || builder[builder.Length - 1] != '-') {
                    builder.Append('-');
                }
            }

            string product = builder.ToString().Trim('-');
            return (string.IsNullOrWhiteSpace(product) ? "OfficeIMO" : product) + "/2.0";
        }

        private void ThrowIfDisposed() {
            if (_disposed) {
                throw new ObjectDisposedException(nameof(GoogleWorkspaceHttpTransport));
            }
        }
    }

    public sealed class GoogleWorkspaceHttpResponse {
        internal GoogleWorkspaceHttpResponse(
            HttpStatusCode statusCode,
            byte[] body,
            string? mediaType,
            IReadOnlyDictionary<string, IReadOnlyList<string>> headers) {
            StatusCode = statusCode;
            Body = body ?? Array.Empty<byte>();
            MediaType = mediaType;
            Headers = headers ?? throw new ArgumentNullException(nameof(headers));
        }

        public HttpStatusCode StatusCode { get; }
        public byte[] Body { get; }
        public string? MediaType { get; }
        public IReadOnlyDictionary<string, IReadOnlyList<string>> Headers { get; }
        public string BodyText => Encoding.UTF8.GetString(Body);

        public string? GetHeader(string name) {
            return Headers.TryGetValue(name, out var values) ? values.FirstOrDefault() : null;
        }

        [RequiresUnreferencedCode("Use DeserializeJson(JsonTypeInfo<T>) in trimmed applications.")]
        [RequiresDynamicCode("Use DeserializeJson(JsonTypeInfo<T>) in NativeAOT applications.")]
        public T DeserializeJson<T>() {
            var value = JsonSerializer.Deserialize<T>(Body, new JsonSerializerOptions {
                PropertyNameCaseInsensitive = true,
            });
            if (value == null) {
                throw new InvalidOperationException("The Google Workspace response body could not be deserialized.");
            }

            return value;
        }

        /// <summary>Deserializes the response body with source-generated JSON metadata.</summary>
        public T DeserializeJson<T>(JsonTypeInfo<T> typeInfo) {
            if (typeInfo == null) throw new ArgumentNullException(nameof(typeInfo));
            var value = JsonSerializer.Deserialize(Body, typeInfo);
            if (value == null) {
                throw new InvalidOperationException("The Google Workspace response body could not be deserialized.");
            }

            return value;
        }
    }

    /// <summary>
    /// Typed failure returned for a non-success Google API response.
    /// </summary>
    public sealed class GoogleWorkspaceApiException : HttpRequestException {
        private GoogleWorkspaceApiException(
            string message,
            string serviceName,
            HttpMethod method,
            string requestUri,
            HttpStatusCode statusCode,
            string responseBody)
            : base(message) {
            ServiceName = serviceName;
            Method = method;
            RequestUri = requestUri;
            ResponseStatusCode = statusCode;
            ResponseBody = responseBody;
        }

        public string ServiceName { get; }
        public HttpMethod Method { get; }
        public string RequestUri { get; }
        public HttpStatusCode ResponseStatusCode { get; }
        public string ResponseBody { get; }

        internal static GoogleWorkspaceApiException Create(
            string serviceName,
            HttpMethod method,
            string requestUri,
            HttpStatusCode statusCode,
            string responseBody) {
            string formattedError = GoogleWorkspaceApiErrorFormatter.Format(responseBody) ?? responseBody;
            string message = $"{serviceName} request to '{requestUri}' failed with {(int)statusCode}: {formattedError}";
            return new GoogleWorkspaceApiException(message, serviceName, method, requestUri, statusCode, responseBody);
        }
    }
}
