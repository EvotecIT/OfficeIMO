using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Dependency-light HTTP transport shared by Google Workspace domain packages.
    /// </summary>
    public sealed class GoogleWorkspaceHttpTransport : IDisposable {
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

        public async Task<TResponse> SendAsync<TResponse>(
            string accessToken,
            HttpMethod method,
            string uri,
            Func<HttpContent?>? contentFactory,
            GoogleWorkspaceRequestSafety requestSafety,
            string serviceName,
            TranslationReport report,
            CancellationToken cancellationToken = default) {
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

                var result = JsonSerializer.Deserialize<TResponse>(body, JsonOptions);
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
            CancellationToken cancellationToken = default) {
            ThrowIfDisposed();
            string effectiveUri = AppendQueryParameter(uri, "quotaUser", _options.QuotaUser);
            string? requestId = _options.RequestIdFactory?.Invoke();
            var retryOptions = GoogleWorkspaceRetryOptions.FromSessionOptions(_options);

            using (var response = await GoogleWorkspaceRetryPolicy.SendAsync(
                _client,
                () => CreateRequest(accessToken, method, effectiveUri, null, requestId),
                retryOptions,
                requestSafety,
                _options.RequestTimeout,
                cancellationToken,
                retryEvent => ReportRetry(report, serviceName, retryEvent)).ConfigureAwait(false)) {
                if (!response.IsSuccessStatusCode) {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    throw GoogleWorkspaceApiException.Create(serviceName, method, effectiveUri, response.StatusCode, body);
                }

                return await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
            }
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
            string? requestId) {
            var request = new HttpRequestMessage(method, uri);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
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

        public T DeserializeJson<T>() {
            var value = JsonSerializer.Deserialize<T>(Body, new JsonSerializerOptions {
                PropertyNameCaseInsensitive = true,
            });
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
