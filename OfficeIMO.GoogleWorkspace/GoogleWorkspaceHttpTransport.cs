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
            CancellationToken cancellationToken = default) {
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
                cancellationToken);
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
            CancellationToken cancellationToken = default) {
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
                cancellationToken);
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
            CancellationToken cancellationToken = default) {
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
                cancellationToken);
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
            CancellationToken cancellationToken = default) {
            return SendAsyncCore(
                accessToken,
                method,
                uri,
                contentFactory,
                requestSafety,
                serviceName,
                report,
                body => JsonSerializer.Deserialize<TResponse>(body, JsonOptions),
                cancellationToken);
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
            CancellationToken cancellationToken = default) {
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
                cancellationToken);
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
            CancellationToken cancellationToken) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(accessToken)) throw new ArgumentException("Access token is required.", nameof(accessToken));
            if (method == null) throw new ArgumentNullException(nameof(method));
            if (string.IsNullOrWhiteSpace(uri)) throw new ArgumentException("Request URI is required.", nameof(uri));
            if (string.IsNullOrWhiteSpace(serviceName)) throw new ArgumentException("Service name is required.", nameof(serviceName));
            if (report == null) throw new ArgumentNullException(nameof(report));

            string effectiveUri = AppendQueryParameter(uri, "quotaUser", _options.QuotaUser);
            string? requestId = _options.RequestIdFactory?.Invoke();
            var retryOptions = GoogleWorkspaceRetryOptions.FromSessionOptions(_options);

            using (var response = await GoogleWorkspaceRetryPolicy.SendAsync(
                _client,
                () => CreateRequest(accessToken, method, effectiveUri, contentFactory, requestId),
                retryOptions,
                requestSafety,
                _options.RequestTimeout,
                cancellationToken,
                retryEvent => ReportRetry(report, serviceName, retryEvent)).ConfigureAwait(false)) {
                string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode) {
                    throw GoogleWorkspaceApiException.Create(serviceName, method, effectiveUri, response.StatusCode, body);
                }

                if (typeof(TResponse) == typeof(object) || string.IsNullOrWhiteSpace(body)) {
                    return default!;
                }

                var result = deserialize(body);
                if (result == null) {
                    throw new InvalidOperationException($"{serviceName} response from '{effectiveUri}' could not be deserialized.");
                }

                return result;
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
            long? maxResponseBytes = null) {
            ThrowIfDisposed();
            if (maxResponseBytes.HasValue && maxResponseBytes.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(maxResponseBytes));
            }
            string effectiveUri = preserveRequestUri
                ? uri
                : AppendQueryParameter(uri, "quotaUser", _options.QuotaUser);
            string? requestId = _options.RequestIdFactory?.Invoke();
            var retryOptions = GoogleWorkspaceRetryOptions.FromSessionOptions(_options);

            return await GoogleWorkspaceRetryPolicy.SendAndProcessAsync(
                _client,
                () => CreateRequest(accessToken, method, effectiveUri, null, requestId, includeAuthorization),
                retryOptions,
                requestSafety,
                _options.RequestTimeout,
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
                            method, effectiveUri, response.StatusCode, body);
                    }

                    return await ReadResponseBytesAsync(response.Content,
                        maxResponseBytes, responseToken).ConfigureAwait(false);
                },
                retryEvent => ReportRetry(report, serviceName, retryEvent))
                .ConfigureAwait(false);
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
            bool preserveRequestUri = false) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(accessToken)) throw new ArgumentException("Access token is required.", nameof(accessToken));
            if (method == null) throw new ArgumentNullException(nameof(method));
            if (string.IsNullOrWhiteSpace(uri)) throw new ArgumentException("Request URI is required.", nameof(uri));
            if (string.IsNullOrWhiteSpace(serviceName)) throw new ArgumentException("Service name is required.", nameof(serviceName));
            if (report == null) throw new ArgumentNullException(nameof(report));

            string effectiveUri = preserveRequestUri
                ? uri
                : AppendQueryParameter(uri, "quotaUser", _options.QuotaUser);
            string? requestId = _options.RequestIdFactory?.Invoke();
            var retryOptions = GoogleWorkspaceRetryOptions.FromSessionOptions(_options);

            using (var response = await GoogleWorkspaceRetryPolicy.SendAsync(
                _client,
                () => {
                    var request = CreateRequest(accessToken, method, effectiveUri, contentFactory, requestId);
                    configureRequest?.Invoke(request);
                    return request;
                },
                retryOptions,
                requestSafety,
                _options.RequestTimeout,
                cancellationToken,
                retryEvent => ReportRetry(report, serviceName, retryEvent)).ConfigureAwait(false)) {
                byte[] body = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                bool accepted = response.IsSuccessStatusCode
                    || (additionalSuccessStatusCodes != null && additionalSuccessStatusCodes.Contains(response.StatusCode));
                if (!accepted) {
                    string responseText = Encoding.UTF8.GetString(body);
                    throw GoogleWorkspaceApiException.Create(serviceName, method, effectiveUri, response.StatusCode, responseText);
                }

                var headers = response.Headers
                    .Concat(response.Content.Headers)
                    .GroupBy(header => header.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(
                        group => group.Key,
                        group => (IReadOnlyList<string>)group.SelectMany(header => header.Value).ToArray(),
                        StringComparer.OrdinalIgnoreCase);
                return new GoogleWorkspaceHttpResponse(
                    response.StatusCode,
                    body,
                    response.Content.Headers.ContentType?.MediaType,
                    headers);
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

        private void ReportRetry(TranslationReport report, string serviceName, GoogleWorkspaceRetryEvent retryEvent) {
            GoogleWorkspaceDiagnosticsDispatcher.AddUnique(
                report,
                _options,
                TranslationSeverity.Info,
                "ApiRetries",
                $"{serviceName} retried {retryEvent.Method} {retryEvent.Uri} after transient {retryEvent.Trigger} using {retryEvent.DelayStrategy} ({retryEvent.Delay.TotalMilliseconds:0} ms, retry {retryEvent.RetryAttempt} of {retryEvent.MaxRetryCount}).",
                $"{retryEvent.Method} {retryEvent.Uri}",
                code: GoogleWorkspaceDiagnosticCodes.ApiRetry);
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
