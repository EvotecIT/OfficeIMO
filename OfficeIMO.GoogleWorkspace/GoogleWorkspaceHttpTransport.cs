using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Net;
using System.Net.Http.Headers;
using System.Runtime.ExceptionServices;
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
        private static readonly HashSet<string> OperationDefiningQueryParameters = new HashSet<string>(
            new[] {
                "addParents",
                "enforceSingleParent",
                "keepRevisionForever",
                "moveToNewOwnersRoot",
                "removeParents",
                "sendNotificationEmail",
                "supportsAllDrives",
                "transferOwnership",
                "uploadType",
                "useDomainAdminAccess",
            }, StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<string> SensitiveQueryParameters = new HashSet<string>(
            new[] {
                "access_token",
                "emailMessage",
                "key",
                "oauth_token",
                "pageToken",
                "quotaUser",
                "startPageToken",
                "token",
                "upload_id",
            }, StringComparer.OrdinalIgnoreCase);
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions {
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = null,
            WriteIndented = false,
        };

        private readonly GoogleWorkspaceSessionOptions _options;
        private readonly HttpClient _client;
        private readonly bool _ownsClient;
        private readonly Func<string, IReadOnlyCollection<string>, string>? _mutationCredentialVerifier;
        private bool _disposed;

        public GoogleWorkspaceHttpTransport(GoogleWorkspaceSessionOptions options)
            : this(options, null) { }

        /// <summary>Creates a transport whose mutations are bound to credentials acquired by the supplied session.</summary>
        public GoogleWorkspaceHttpTransport(GoogleWorkspaceSession session)
            : this(session?.Options ?? throw new ArgumentNullException(nameof(session)),
                session.VerifyMutationCredential) { }

        private GoogleWorkspaceHttpTransport(GoogleWorkspaceSessionOptions options,
            Func<string, IReadOnlyCollection<string>, string>? mutationCredentialVerifier) {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _mutationCredentialVerifier = mutationCredentialVerifier;
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
            MutationAttempt? mutation = BeginMutation(accessToken, method, visibleTarget, requestSafety, mutationKind,
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

                    // The remote mutation has already committed before its response is read or decoded.
                    // Persist that outcome first so response-processing failures cannot invite a blind retry.
                    CompleteMutationSuccess(mutation);
                    try {
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
                    } catch (Exception exception) when (mutation != null) {
                        throw new AcceptedMutationResponseProcessingException(exception);
                    }
                    },
                    retryEvent => { mutation?.CountRetry(); ReportRetry(report, serviceName, retryEvent, visibleTarget); })
                    .ConfigureAwait(false);
                CompleteMutationSuccess(mutation);
                return result;
            } catch (AcceptedMutationResponseProcessingException exception) {
                CompleteMutationFailure(mutation, exception.InnerException!);
                ExceptionDispatchInfo.Capture(exception.InnerException!).Throw();
                throw;
            } catch (GoogleWorkspaceNoResponseException exception) when (mutation != null) {
                throw CompleteMutationAmbiguous(mutation, exception.InnerException!);
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
            MutationAttempt? mutation = BeginMutation(accessToken, method, visibleTarget, requestSafety, mutationKind,
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

                    CompleteMutationSuccess(mutation);
                    try {
                        validateResponse?.Invoke(response);
                        return await ReadResponseBytesAsync(response.Content,
                            maxResponseBytes, responseToken).ConfigureAwait(false);
                    } catch (Exception exception) when (mutation != null) {
                        throw new AcceptedMutationResponseProcessingException(exception);
                    }
                    },
                    retryEvent => { mutation?.CountRetry(); ReportRetry(report, serviceName, retryEvent, visibleTarget); })
                    .ConfigureAwait(false);
                CompleteMutationSuccess(mutation);
                return result;
            } catch (AcceptedMutationResponseProcessingException exception) {
                CompleteMutationFailure(mutation, exception.InnerException!);
                ExceptionDispatchInfo.Capture(exception.InnerException!).Throw();
                throw;
            } catch (GoogleWorkspaceNoResponseException exception) when (mutation != null) {
                throw CompleteMutationAmbiguous(mutation, exception.InnerException!);
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
            if (content.Headers.ContentLength is long contentLength &&
                contentLength >= 0 &&
                contentLength <= int.MaxValue) {
                int targetLength = checked((int)(truncateAtLimit && limit.HasValue
                    ? Math.Min(contentLength, limit.Value)
                    : contentLength));
                return await ReadDeclaredLengthResponseAsync(
                    input,
                    targetLength,
                    limit,
                    cancellationToken,
                    truncateAtLimit).ConfigureAwait(false);
            }

            using var output = new MemoryStream();
            byte[] buffer = ArrayPool<byte>.Shared.Rent(81920);
            try {
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
            } finally {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        private static async Task<byte[]> ReadDeclaredLengthResponseAsync(
            Stream input,
            int targetLength,
            long? limit,
            CancellationToken cancellationToken,
            bool truncateAtLimit) {
            byte[] result = new byte[targetLength];
            int total = 0;
            while (total < result.Length) {
                int read = await input.ReadAsync(
                    result,
                    total,
                    result.Length - total,
                    cancellationToken).ConfigureAwait(false);
                if (read == 0) {
                    if (total != result.Length) {
                        Array.Resize(ref result, total);
                    }
                    return result;
                }
                total += read;
            }

            if (truncateAtLimit && limit.HasValue && total == limit.Value) {
                return result;
            }

            byte[] probe = ArrayPool<byte>.Shared.Rent(1);
            try {
                int extra = await input.ReadAsync(probe, 0, 1, cancellationToken).ConfigureAwait(false);
                if (extra == 0) {
                    return result;
                }
                if (limit.HasValue && total >= limit.Value) {
                    throw new InvalidDataException(
                        $"The response exceeded the configured limit of {limit.Value} bytes.");
                }

                using var output = new MemoryStream(checked(result.Length + 81920));
                output.Write(result, 0, result.Length);
                output.Write(probe, 0, extra);
                byte[] buffer = ArrayPool<byte>.Shared.Rent(81920);
                try {
                    long copied = total + extra;
                    while (true) {
                        int read = await input.ReadAsync(buffer, 0, buffer.Length, cancellationToken)
                            .ConfigureAwait(false);
                        if (read == 0) break;
                        if (limit.HasValue && read > limit.Value - copied) {
                            throw new InvalidDataException(
                                $"The response exceeded the configured limit of {limit.Value} bytes.");
                        }
                        output.Write(buffer, 0, read);
                        copied += read;
                    }
                    return output.ToArray();
                } finally {
                    ArrayPool<byte>.Shared.Return(buffer);
                }
            } finally {
                ArrayPool<byte>.Shared.Return(probe);
            }
        }

        public Task<GoogleWorkspaceHttpResponse> SendRawAsync(
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
            return SendRawAsyncCore(accessToken, method, uri, contentFactory, requestSafety, serviceName,
                report, cancellationToken, configureRequest, additionalSuccessStatusCodes, preserveRequestUri,
                diagnosticTarget, mutationKind, revisionPrecondition, requiredScopes, potentialDataLoss,
                allowSafeNonReadMethod: false);
        }

        internal Task<GoogleWorkspaceHttpResponse> SendRawSafeProbeAsync(
            string accessToken,
            HttpMethod method,
            string uri,
            Func<HttpContent?> contentFactory,
            string serviceName,
            TranslationReport report,
            CancellationToken cancellationToken,
            IReadOnlyCollection<HttpStatusCode>? additionalSuccessStatusCodes = null,
            bool preserveRequestUri = false,
            string? diagnosticTarget = null) {
            return SendRawAsyncCore(accessToken, method, uri, contentFactory,
                GoogleWorkspaceRequestSafety.Safe, serviceName, report, cancellationToken, null,
                additionalSuccessStatusCodes, preserveRequestUri, diagnosticTarget,
                GoogleWorkspaceMutationKind.Unspecified, null, null, false,
                allowSafeNonReadMethod: true);
        }

        private async Task<GoogleWorkspaceHttpResponse> SendRawAsyncCore(
            string accessToken,
            HttpMethod method,
            string uri,
            Func<HttpContent?>? contentFactory,
            GoogleWorkspaceRequestSafety requestSafety,
            string serviceName,
            TranslationReport report,
            CancellationToken cancellationToken,
            Action<HttpRequestMessage>? configureRequest,
            IReadOnlyCollection<HttpStatusCode>? additionalSuccessStatusCodes,
            bool preserveRequestUri,
            string? diagnosticTarget,
            GoogleWorkspaceMutationKind mutationKind,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition,
            IReadOnlyCollection<string>? requiredScopes,
            bool potentialDataLoss,
            bool allowSafeNonReadMethod) {
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
            MutationAttempt? mutation = BeginMutation(accessToken, method, visibleTarget, requestSafety, mutationKind,
                revisionPrecondition, serviceName, requestId, requiredScopes, retryOptions, potentialDataLoss,
                allowSafeNonReadMethod);

            try {
                GoogleWorkspaceHttpResponse result = await GoogleWorkspaceRetryPolicy.SendAndProcessAsync(
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
                    async (response, responseToken) => {
                    bool accepted = response.IsSuccessStatusCode
                        || (additionalSuccessStatusCodes != null && additionalSuccessStatusCodes.Contains(response.StatusCode));
                    if (!accepted) {
                        byte[] errorBody = await ReadResponseBytesAsync(response.Content,
                            MaximumErrorResponseBytes, responseToken, truncateAtLimit: true).ConfigureAwait(false);
                        string responseText = Encoding.UTF8.GetString(errorBody);
                        throw GoogleWorkspaceApiException.Create(serviceName, method, visibleTarget, response.StatusCode, responseText);
                    }

                    CompleteMutationSuccess(mutation);
                    try {
                        byte[] body = await ReadResponseBytesAsync(response.Content, null,
                            responseToken).ConfigureAwait(false);
                        var headers = response.Headers
                            .Concat(response.Content.Headers)
                            .GroupBy(header => header.Key, StringComparer.OrdinalIgnoreCase)
                            .ToDictionary(group => group.Key,
                                group => (IReadOnlyList<string>)group.SelectMany(header => header.Value).ToArray(),
                                StringComparer.OrdinalIgnoreCase);
                        return new GoogleWorkspaceHttpResponse(response.StatusCode, body,
                            response.Content.Headers.ContentType?.MediaType, headers);
                    } catch (Exception exception) when (mutation != null) {
                        throw new AcceptedMutationResponseProcessingException(exception);
                    }
                    },
                    retryEvent => { mutation?.CountRetry(); ReportRetry(report, serviceName, retryEvent, visibleTarget); })
                    .ConfigureAwait(false);
                CompleteMutationSuccess(mutation);
                return result;
            } catch (AcceptedMutationResponseProcessingException exception) {
                CompleteMutationFailure(mutation, exception.InnerException!);
                ExceptionDispatchInfo.Capture(exception.InnerException!).Throw();
                throw;
            } catch (GoogleWorkspaceNoResponseException exception) when (mutation != null) {
                throw CompleteMutationAmbiguous(mutation, exception.InnerException!);
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

        internal DeferredMutation BeginDeferredMutation(string accessToken, HttpMethod method, string target,
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
            MutationAttempt mutation = BeginMutation(accessToken, method, target, requestSafety, mutationKind,
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

        private MutationAttempt? BeginMutation(string accessToken, HttpMethod method, string target,
            GoogleWorkspaceRequestSafety requestSafety, GoogleWorkspaceMutationKind mutationKind,
            GoogleWorkspaceRevisionPrecondition? revisionPrecondition, string service, string? requestId,
            IReadOnlyCollection<string>? requiredScopes, GoogleWorkspaceRetryOptions retryOptions,
            bool adapterDeclaredPotentialDataLoss,
            bool allowSafeNonReadMethod = false) {
            if (requestSafety == GoogleWorkspaceRequestSafety.Safe) {
                bool safeMethod = method == HttpMethod.Get || method == HttpMethod.Head
                    || method == HttpMethod.Options;
                if ((!safeMethod && !allowSafeNonReadMethod)
                    || mutationKind != GoogleWorkspaceMutationKind.Unspecified
                    || revisionPrecondition != null || adapterDeclaredPotentialDataLoss) {
                    throw new InvalidOperationException(
                        "Safe request semantics are valid only for GET, HEAD, or OPTIONS operations without mutation declarations.");
                }
                return null;
            }
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
            if (_mutationCredentialVerifier == null) {
                throw new InvalidOperationException("Google mutations require a transport bound to the GoogleWorkspaceSession that acquired the access token.");
            }
            string credentialAccount = _mutationCredentialVerifier(accessToken, Array.AsReadOnly(scopeSnapshot));
            if (!StringComparer.OrdinalIgnoreCase.Equals(credentialAccount, expectedAccount)) {
                throw new InvalidOperationException("The acquired Google credential account does not match the configured session account.");
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

        private sealed class AcceptedMutationResponseProcessingException : Exception {
            internal AcceptedMutationResponseProcessingException(Exception innerException)
                : base("An accepted mutation response could not be processed.", innerException) { }
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

        private static GoogleWorkspaceAmbiguousMutationException CompleteMutationAmbiguous(
            MutationAttempt mutation, Exception transportFailure) {
            GoogleWorkspaceReceiptPersistenceException? persistenceFailure =
                mutation.Complete(false, "ambiguous-no-response", out GoogleWorkspaceOperationReceipt? receipt,
                    isOutcomeAmbiguous: true);
            GoogleWorkspaceOperationReceipt operationReceipt = receipt
                ?? throw new InvalidOperationException("The ambiguous mutation receipt was already completed.");
            var exception = new GoogleWorkspaceAmbiguousMutationException(operationReceipt, transportFailure);
            if (persistenceFailure != null) {
                exception.Data[GoogleWorkspaceReceiptPersistenceException.ExceptionDataKey] = persistenceFailure;
            }
            return exception;
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
                return Complete(succeeded, outcome, out _);
            }
            internal GoogleWorkspaceReceiptPersistenceException? Complete(bool succeeded, string outcome,
                out GoogleWorkspaceOperationReceipt? receipt, bool isOutcomeAmbiguous = false) {
                receipt = null;
                if (_completed) return null;
                _completed = true;
                string? enforcedRevision = _context.RevisionPreconditionKind == GoogleWorkspaceRevisionPreconditionKind.HttpEntityTag
                    ? _policy.ExpectedRevision
                    : _context.AdapterExpectedRevision;
                var operationReceipt = new GoogleWorkspaceOperationReceipt(_policy, _context.Service, _context.Method,
                    _context.Target, _context.RequestId, _retryCount, succeeded, outcome,
                    _context.MutationKind, _context.RevisionPreconditionKind, enforcedRevision,
                    isOutcomeAmbiguous);
                receipt = operationReceipt;
                try {
                    _sink(operationReceipt);
                    return null;
                } catch (Exception exception) {
                    return new GoogleWorkspaceReceiptPersistenceException(operationReceipt, succeeded, exception);
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
            int fragment = uri.IndexOf('#');
            string withoutFragment = fragment < 0 ? uri : uri.Substring(0, fragment);
            int query = withoutFragment.IndexOf('?');
            if (query < 0) {
                return withoutFragment;
            }

            string target = withoutFragment.Substring(0, query);
            string queryText = withoutFragment.Substring(query + 1);
            var visibleParameters = new List<string>();
            foreach (string segment in queryText.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries)) {
                int equals = segment.IndexOf('=');
                string rawName = equals < 0 ? segment : segment.Substring(0, equals);
                string name;
                try {
                    name = Uri.UnescapeDataString(rawName.Replace("+", " "));
                } catch (UriFormatException) {
                    continue;
                }

                if (SensitiveQueryParameters.Contains(name)) {
                    visibleParameters.Add(Uri.EscapeDataString(name) + "=%3Credacted%3E");
                } else if (OperationDefiningQueryParameters.Contains(name)) {
                    visibleParameters.Add(equals < 0
                        ? Uri.EscapeDataString(name)
                        : Uri.EscapeDataString(name) + "=" + segment.Substring(equals + 1));
                }
            }

            return visibleParameters.Count == 0
                ? target
                : target + "?" + string.Join("&", visibleParameters);
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
