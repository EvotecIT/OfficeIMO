using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.IO;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Describes whether a request can be repeated after an ambiguous transport outcome.
    /// </summary>
    public enum GoogleWorkspaceRequestSafety {
        /// <summary>Reading or otherwise side-effect-free operation.</summary>
        Safe = 0,
        /// <summary>Mutation whose repeated application has the same intended outcome.</summary>
        Idempotent = 1,
        /// <summary>Mutation that can create duplicates or apply an action more than once.</summary>
        NonIdempotent = 2,
    }

    internal sealed class GoogleWorkspaceNoResponseException : Exception {
        internal GoogleWorkspaceNoResponseException(Exception innerException)
            : base("The request ended before response headers were received.", innerException) { }
    }

    public sealed class GoogleWorkspaceRetryOptions {
        public GoogleWorkspaceRetryOptions(int maxRetryCount, TimeSpan baseDelay, TimeSpan maxDelay,
            GoogleWorkspaceSessionOptions? sessionOptions = null) {
            MaxRetryCount = Math.Max(0, maxRetryCount);
            BaseDelay = baseDelay <= TimeSpan.Zero ? TimeSpan.FromMilliseconds(200) : baseDelay;
            MaxDelay = maxDelay <= TimeSpan.Zero ? TimeSpan.FromSeconds(5) : maxDelay;
            SessionOptions = sessionOptions;
            MaxElapsedTime = sessionOptions?.MaxRetryElapsedTime ?? TimeSpan.FromMinutes(2);
            RateLimitPolicy = sessionOptions?.RateLimitPolicy ?? GoogleWorkspaceRateLimitPolicy.HonorRetryAfter;
            if (MaxDelay < BaseDelay) {
                MaxDelay = BaseDelay;
            }
        }

        public int MaxRetryCount { get; }
        public TimeSpan BaseDelay { get; }
        public TimeSpan MaxDelay { get; }
        public GoogleWorkspaceSessionOptions? SessionOptions { get; }
        public TimeSpan MaxElapsedTime { get; }
        public GoogleWorkspaceRateLimitPolicy RateLimitPolicy { get; }

        public static GoogleWorkspaceRetryOptions FromSessionOptions(GoogleWorkspaceSessionOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            return new GoogleWorkspaceRetryOptions(options.MaxRetryCount, options.RetryBaseDelay, options.RetryMaxDelay, options);
        }
    }

    public sealed class GoogleWorkspaceRetryEvent {
        public GoogleWorkspaceRetryEvent(
            string method,
            string uri,
            int retryAttempt,
            int maxRetryCount,
            string trigger,
            TimeSpan delay,
            string delayStrategy) {
            Method = method ?? string.Empty;
            Uri = uri ?? string.Empty;
            RetryAttempt = retryAttempt;
            MaxRetryCount = maxRetryCount;
            Trigger = trigger ?? string.Empty;
            Delay = delay;
            DelayStrategy = delayStrategy ?? string.Empty;
        }

        public string Method { get; }
        public string Uri { get; }
        public int RetryAttempt { get; }
        public int MaxRetryCount { get; }
        public string Trigger { get; }
        public TimeSpan Delay { get; }
        public string DelayStrategy { get; }
    }

    public static class GoogleWorkspaceRetryPolicy {
        private const int MaximumRateLimitErrorBytes = 64 * 1024;

        public static Task<HttpResponseMessage> SendAsync(
            HttpClient client,
            Func<HttpRequestMessage> requestFactory,
            GoogleWorkspaceRetryOptions retryOptions,
            GoogleWorkspaceRequestSafety requestSafety,
            CancellationToken cancellationToken,
            Action<GoogleWorkspaceRetryEvent>? onRetry = null) {
            return SendAsync(
                client,
                requestFactory,
                retryOptions,
                requestSafety,
                Timeout.InfiniteTimeSpan,
                cancellationToken,
                onRetry);
        }

        public static async Task<HttpResponseMessage> SendAsync(
            HttpClient client,
            Func<HttpRequestMessage> requestFactory,
            GoogleWorkspaceRetryOptions retryOptions,
            GoogleWorkspaceRequestSafety requestSafety,
            TimeSpan requestTimeout,
            CancellationToken cancellationToken,
            Action<GoogleWorkspaceRetryEvent>? onRetry = null) {
            return await SendCoreAsync(
                client,
                requestFactory,
                retryOptions,
                requestSafety,
                requestTimeout,
                HttpCompletionOption.ResponseContentRead,
                cancellationToken,
                (response, _) => Task.FromResult(response),
                disposeFinalResponse: false,
                wrapMutationNoResponse: false,
                onRetry).ConfigureAwait(false);
        }

        internal static Task<TResult> SendAndProcessAsync<TResult>(
            HttpClient client,
            Func<HttpRequestMessage> requestFactory,
            GoogleWorkspaceRetryOptions retryOptions,
            GoogleWorkspaceRequestSafety requestSafety,
            TimeSpan requestTimeout,
            CancellationToken cancellationToken,
            Func<HttpResponseMessage, CancellationToken, Task<TResult>>
                responseHandler,
            Action<GoogleWorkspaceRetryEvent>? onRetry = null) {
            if (responseHandler == null) {
                throw new ArgumentNullException(nameof(responseHandler));
            }
            return SendCoreAsync(
                client,
                requestFactory,
                retryOptions,
                requestSafety,
                requestTimeout,
                HttpCompletionOption.ResponseHeadersRead,
                cancellationToken,
                responseHandler,
                disposeFinalResponse: true,
                wrapMutationNoResponse: true,
                onRetry);
        }

        private static async Task<TResult> SendCoreAsync<TResult>(
            HttpClient client,
            Func<HttpRequestMessage> requestFactory,
            GoogleWorkspaceRetryOptions retryOptions,
            GoogleWorkspaceRequestSafety requestSafety,
            TimeSpan requestTimeout,
            HttpCompletionOption completionOption,
            CancellationToken cancellationToken,
            Func<HttpResponseMessage, CancellationToken, Task<TResult>>
                responseHandler,
            bool disposeFinalResponse,
            bool wrapMutationNoResponse,
            Action<GoogleWorkspaceRetryEvent>? onRetry) {
            if (retryOptions == null) throw new ArgumentNullException(nameof(retryOptions));
            int retryBudget = retryOptions.MaxRetryCount;
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            using var deadlineSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            deadlineSource.CancelAfter(retryOptions.MaxElapsedTime);

            for (int attempt = 0; ; attempt++) {
                EnsureElapsedBudget(stopwatch.Elapsed, TimeSpan.Zero, retryOptions);
                using (var timeoutSource = CancellationTokenSource.CreateLinkedTokenSource(deadlineSource.Token))
                using (var request = requestFactory()) {
                    if (requestTimeout > TimeSpan.Zero && requestTimeout != Timeout.InfiniteTimeSpan) {
                        timeoutSource.CancelAfter(requestTimeout);
                    }

                    string method = request.Method.Method;
                    string uri = request.RequestUri?.AbsoluteUri ?? string.Empty;

                    HttpResponseMessage response;
                    try {
                        response = await client.SendAsync(request,
                            completionOption,
                            timeoutSource.Token).ConfigureAwait(false);
                    } catch (TaskCanceledException exception) when (!cancellationToken.IsCancellationRequested && deadlineSource.IsCancellationRequested) {
                        var timeout = new TimeoutException(
                            "The configured Google Workspace retry elapsed-time budget was exhausted.", exception);
                        if (wrapMutationNoResponse
                            && requestSafety != GoogleWorkspaceRequestSafety.Safe) {
                            throw new GoogleWorkspaceNoResponseException(timeout);
                        }
                        throw timeout;
                    } catch (Exception exception) when (wrapMutationNoResponse
                        && requestSafety != GoogleWorkspaceRequestSafety.Safe
                        && IsNoResponseFailure(exception)) {
                        // A guarded mutation may already have committed. Even an idempotent retry cannot
                        // prove the first attempt's audit outcome, and a desired-state response such as 404
                        // can erase that uncertainty. Require reconciliation before any caller retry.
                        throw new GoogleWorkspaceNoResponseException(exception);
                    } catch (HttpRequestException exception) when (!(exception is GoogleWorkspaceApiException)
                        && CanRetry(requestSafety) && attempt < retryBudget) {
                        var (delay, delayStrategy) = GetRetryDelay(null, attempt, retryOptions);
                        EnsureElapsedBudget(stopwatch.Elapsed, delay, retryOptions);
                        onRetry?.Invoke(new GoogleWorkspaceRetryEvent(
                            method,
                            uri,
                            attempt + 1,
                            retryBudget,
                            "network failure",
                            delay,
                            delayStrategy));
                        await DelayWithinDeadlineAsync(delay, deadlineSource.Token, cancellationToken).ConfigureAwait(false);
                        continue;
                    } catch (TaskCanceledException) when (!cancellationToken.IsCancellationRequested && CanRetry(requestSafety) && attempt < retryBudget) {
                        var (delay, delayStrategy) = GetRetryDelay(null, attempt, retryOptions);
                        EnsureElapsedBudget(stopwatch.Elapsed, delay, retryOptions);
                        onRetry?.Invoke(new GoogleWorkspaceRetryEvent(
                            method,
                            uri,
                            attempt + 1,
                            retryBudget,
                            "request timeout",
                            delay,
                            delayStrategy));
                        await DelayWithinDeadlineAsync(delay, deadlineSource.Token, cancellationToken).ConfigureAwait(false);
                        continue;
                    }

                    try {
                        bool retryableResponse = CanRetry(requestSafety)
                            && attempt < retryBudget
                            && await ShouldRetryAsync(response, retryOptions.RateLimitPolicy,
                                timeoutSource.Token).ConfigureAwait(false);
                        if (!retryableResponse) {
                            if (!disposeFinalResponse) {
                                TResult result = await responseHandler(response,
                                    timeoutSource.Token).ConfigureAwait(false);
                                try {
                                    EnsureElapsedBudget(stopwatch.Elapsed, TimeSpan.Zero, retryOptions);
                                } catch {
                                    response.Dispose();
                                    throw;
                                }
                                return result;
                            }
                            using (response) {
                                TResult result = await responseHandler(response,
                                    timeoutSource.Token).ConfigureAwait(false);
                                EnsureElapsedBudget(stopwatch.Elapsed, TimeSpan.Zero, retryOptions);
                                return result;
                            }
                        }
                    } catch (HttpRequestException exception) when (!(exception is GoogleWorkspaceApiException)
                        && CanRetry(requestSafety) && attempt < retryBudget) {
                        response.Dispose();
                        await DelayAfterTransportFailureAsync(method, uri, attempt, retryBudget,
                            retryOptions, stopwatch.Elapsed, "network failure while reading response", deadlineSource.Token, cancellationToken,
                            onRetry).ConfigureAwait(false);
                        continue;
                    } catch (IOException) when (CanRetry(requestSafety) && attempt < retryBudget) {
                        response.Dispose();
                        await DelayAfterTransportFailureAsync(method, uri, attempt, retryBudget,
                            retryOptions, stopwatch.Elapsed, "network failure while reading response", deadlineSource.Token, cancellationToken,
                            onRetry).ConfigureAwait(false);
                        continue;
                    } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested && deadlineSource.IsCancellationRequested) {
                        response.Dispose();
                        throw new TimeoutException("The configured Google Workspace retry elapsed-time budget was exhausted.");
                    } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested && CanRetry(requestSafety) && attempt < retryBudget) {
                        response.Dispose();
                        await DelayAfterTransportFailureAsync(method, uri, attempt, retryBudget,
                            retryOptions, stopwatch.Elapsed, "request timeout while reading response", deadlineSource.Token, cancellationToken,
                            onRetry).ConfigureAwait(false);
                        continue;
                    } catch {
                        response.Dispose();
                        throw;
                    }

                    var (statusDelay, statusDelayStrategy) = GetRetryDelay(
                        response.Headers.RetryAfter, attempt, retryOptions);
                    response.Dispose();
                    EnsureElapsedBudget(stopwatch.Elapsed, statusDelay, retryOptions);
                    onRetry?.Invoke(new GoogleWorkspaceRetryEvent(
                        method,
                        uri,
                        attempt + 1,
                        retryBudget,
                        $"HTTP {(int)response.StatusCode}",
                        statusDelay,
                        statusDelayStrategy));
                    await DelayWithinDeadlineAsync(statusDelay, deadlineSource.Token, cancellationToken)
                        .ConfigureAwait(false);
                }
            }
        }

        private static async Task DelayAfterTransportFailureAsync(
            string method,
            string uri,
            int attempt,
            int retryBudget,
            GoogleWorkspaceRetryOptions retryOptions,
            TimeSpan elapsed,
            string trigger,
            CancellationToken deadlineToken,
            CancellationToken cancellationToken,
            Action<GoogleWorkspaceRetryEvent>? onRetry) {
            var (delay, delayStrategy) = GetRetryDelay(null, attempt, retryOptions);
            EnsureElapsedBudget(elapsed, delay, retryOptions);
            onRetry?.Invoke(new GoogleWorkspaceRetryEvent(
                method,
                uri,
                attempt + 1,
                retryBudget,
                trigger,
                delay,
                delayStrategy));
            await DelayWithinDeadlineAsync(delay, deadlineToken, cancellationToken).ConfigureAwait(false);
        }

        private static async Task DelayWithinDeadlineAsync(TimeSpan delay,
            CancellationToken deadlineToken, CancellationToken cancellationToken) {
            try {
                await Task.Delay(delay, deadlineToken).ConfigureAwait(false);
            } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested) {
                throw new TimeoutException("The configured Google Workspace retry elapsed-time budget was exhausted.");
            }
        }

        private static bool CanRetry(GoogleWorkspaceRequestSafety requestSafety) {
            return requestSafety == GoogleWorkspaceRequestSafety.Safe
                || requestSafety == GoogleWorkspaceRequestSafety.Idempotent;
        }

        private static async Task<bool> ShouldRetryAsync(HttpResponseMessage response,
            GoogleWorkspaceRateLimitPolicy rateLimitPolicy, CancellationToken cancellationToken) {
            if (response.StatusCode != HttpStatusCode.Forbidden) {
                return ShouldRetryStatus(response.StatusCode, rateLimitPolicy);
            }
            if (rateLimitPolicy == GoogleWorkspaceRateLimitPolicy.FailFast) {
                return false;
            }

            byte[] errorBytes = await ReadAndReplaceErrorContentAsync(response,
                MaximumRateLimitErrorBytes, cancellationToken).ConfigureAwait(false);
            try {
                using JsonDocument document = JsonDocument.Parse(errorBytes);
                return ContainsRetryableRateLimitReason(document.RootElement);
            } catch (JsonException) {
                string errorText = Encoding.UTF8.GetString(errorBytes);
                return ContainsRetryableRateLimitReason(errorText);
            }
        }

        private static bool IsNoResponseFailure(Exception exception) =>
            (exception is HttpRequestException && !(exception is GoogleWorkspaceApiException))
            || exception is TimeoutException
            || exception is OperationCanceledException;

        private static bool ContainsRetryableRateLimitReason(JsonElement element) {
            if (element.ValueKind == JsonValueKind.Object) {
                foreach (JsonProperty property in element.EnumerateObject()) {
                    if (StringComparer.OrdinalIgnoreCase.Equals(property.Name, "reason")
                        && property.Value.ValueKind == JsonValueKind.String
                        && ContainsRetryableRateLimitReason(property.Value.GetString() ?? string.Empty)) {
                        return true;
                    }
                    if (ContainsRetryableRateLimitReason(property.Value)) {
                        return true;
                    }
                }
            } else if (element.ValueKind == JsonValueKind.Array) {
                foreach (JsonElement item in element.EnumerateArray()) {
                    if (ContainsRetryableRateLimitReason(item)) {
                        return true;
                    }
                }
            }
            return false;
        }

        private static bool ContainsRetryableRateLimitReason(string reason) =>
            StringComparer.OrdinalIgnoreCase.Equals(reason, "rateLimitExceeded")
            || StringComparer.OrdinalIgnoreCase.Equals(reason, "userRateLimitExceeded")
            || StringComparer.OrdinalIgnoreCase.Equals(reason, "sharingRateLimitExceeded")
            || reason.IndexOf("\"rateLimitExceeded\"", StringComparison.OrdinalIgnoreCase) >= 0
            || reason.IndexOf("\"userRateLimitExceeded\"", StringComparison.OrdinalIgnoreCase) >= 0
            || reason.IndexOf("\"sharingRateLimitExceeded\"", StringComparison.OrdinalIgnoreCase) >= 0;

        private static async Task<byte[]> ReadAndReplaceErrorContentAsync(HttpResponseMessage response,
            int maximumBytes, CancellationToken cancellationToken) {
            HttpContent originalContent = response.Content;
            var originalHeaders = originalContent.Headers
                .Where(header => !StringComparer.OrdinalIgnoreCase.Equals(header.Key, "Content-Length"))
                .ToArray();
            using Stream input = await originalContent.ReadAsStreamAsync().ConfigureAwait(false);
            using var output = new MemoryStream();
            byte[] buffer = new byte[8192];
            while (output.Length < maximumBytes) {
                int count = (int)Math.Min(buffer.Length, maximumBytes - output.Length);
                int read = await input.ReadAsync(buffer, 0, count, cancellationToken).ConfigureAwait(false);
                if (read == 0) break;
                output.Write(buffer, 0, read);
            }
            byte[] bytes = output.ToArray();
            var replacement = new ByteArrayContent(bytes);
            foreach (var header in originalHeaders) {
                replacement.Headers.TryAddWithoutValidation(header.Key, header.Value);
            }
            originalContent.Dispose();
            response.Content = replacement;
            return bytes;
        }

        // Retry only the status codes Google APIs commonly use for throttling or transient infrastructure failures.
        private static bool ShouldRetryStatus(HttpStatusCode statusCode, GoogleWorkspaceRateLimitPolicy rateLimitPolicy) {
            if ((int)statusCode == 429 && rateLimitPolicy == GoogleWorkspaceRateLimitPolicy.FailFast) return false;
            switch (statusCode) {
                case HttpStatusCode.RequestTimeout:
                case (HttpStatusCode)429:
                case HttpStatusCode.InternalServerError:
                case HttpStatusCode.BadGateway:
                case HttpStatusCode.ServiceUnavailable:
                case HttpStatusCode.GatewayTimeout:
                    return true;
                default:
                    return false;
            }
        }

        private static void EnsureElapsedBudget(TimeSpan elapsed, TimeSpan nextDelay,
            GoogleWorkspaceRetryOptions options) {
            if (options.MaxElapsedTime <= TimeSpan.Zero || elapsed + nextDelay > options.MaxElapsedTime) {
                throw new TimeoutException("The configured Google Workspace retry elapsed-time budget was exhausted.");
            }
        }

        private static (TimeSpan Delay, string Strategy) GetRetryDelay(RetryConditionHeaderValue? retryAfter, int retryAttempt, GoogleWorkspaceRetryOptions retryOptions) {
            if (retryAfter?.Delta is TimeSpan retryDelta && retryDelta > TimeSpan.Zero) {
                return (ClampDelay(retryDelta, retryOptions), "server Retry-After");
            }

            if (retryAfter?.Date is DateTimeOffset retryDate) {
                var retryDelay = retryDate - DateTimeOffset.UtcNow;
                if (retryDelay > TimeSpan.Zero) {
                    return (ClampDelay(retryDelay, retryOptions), "server Retry-After");
                }
            }

            int boundedAttempt = Math.Min(retryAttempt, 4);
            double jitter = GetJitterFactor();
            var computedDelay = TimeSpan.FromMilliseconds(retryOptions.BaseDelay.TotalMilliseconds * Math.Pow(2, boundedAttempt) * jitter);
            return (ClampDelay(computedDelay, retryOptions), "exponential backoff");
        }

        private static double GetJitterFactor() {
            byte[] bytes = new byte[4];
            using (var random = RandomNumberGenerator.Create()) {
                random.GetBytes(bytes);
            }

            uint value = BitConverter.ToUInt32(bytes, 0);
            return 0.9d + ((double)value / uint.MaxValue) * 0.2d;
        }

        private static TimeSpan ClampDelay(TimeSpan delay, GoogleWorkspaceRetryOptions retryOptions) {
            if (delay <= TimeSpan.Zero) {
                return retryOptions.BaseDelay;
            }

            return delay > retryOptions.MaxDelay ? retryOptions.MaxDelay : delay;
        }
    }
}
