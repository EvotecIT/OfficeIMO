using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.IO;

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

    public sealed class GoogleWorkspaceRetryOptions {
        public GoogleWorkspaceRetryOptions(int maxRetryCount, TimeSpan baseDelay, TimeSpan maxDelay, GoogleWorkspaceSessionOptions? sessionOptions = null) {
            MaxRetryCount = Math.Max(0, maxRetryCount);
            BaseDelay = baseDelay <= TimeSpan.Zero ? TimeSpan.FromMilliseconds(200) : baseDelay;
            MaxDelay = maxDelay <= TimeSpan.Zero ? TimeSpan.FromSeconds(5) : maxDelay;
            SessionOptions = sessionOptions;
            if (MaxDelay < BaseDelay) {
                MaxDelay = BaseDelay;
            }
        }

        public int MaxRetryCount { get; }
        public TimeSpan BaseDelay { get; }
        public TimeSpan MaxDelay { get; }
        public GoogleWorkspaceSessionOptions? SessionOptions { get; }

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
            Action<GoogleWorkspaceRetryEvent>? onRetry) {
            if (retryOptions == null) throw new ArgumentNullException(nameof(retryOptions));
            int retryBudget = retryOptions.MaxRetryCount;

            for (int attempt = 0; ; attempt++) {
                using (var timeoutSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
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
                    } catch (HttpRequestException) when (CanRetry(requestSafety) && attempt < retryBudget) {
                        var (delay, delayStrategy) = GetRetryDelay(null, attempt, retryOptions);
                        onRetry?.Invoke(new GoogleWorkspaceRetryEvent(
                            method,
                            uri,
                            attempt + 1,
                            retryBudget,
                            "network failure",
                            delay,
                            delayStrategy));
                        await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
                        continue;
                    } catch (TaskCanceledException) when (!cancellationToken.IsCancellationRequested && CanRetry(requestSafety) && attempt < retryBudget) {
                        var (delay, delayStrategy) = GetRetryDelay(null, attempt, retryOptions);
                        onRetry?.Invoke(new GoogleWorkspaceRetryEvent(
                            method,
                            uri,
                            attempt + 1,
                            retryBudget,
                            "request timeout",
                            delay,
                            delayStrategy));
                        await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
                        continue;
                    }

                    if (!CanRetry(requestSafety)
                        || !ShouldRetry(response.StatusCode)
                        || attempt >= retryBudget) {
                        try {
                            if (!disposeFinalResponse) {
                                return await responseHandler(response,
                                    timeoutSource.Token).ConfigureAwait(false);
                            }
                            using (response) {
                                return await responseHandler(response,
                                    timeoutSource.Token).ConfigureAwait(false);
                            }
                        } catch (HttpRequestException) when (CanRetry(requestSafety) && attempt < retryBudget) {
                            response.Dispose();
                            await DelayAfterTransportFailureAsync(method, uri, attempt, retryBudget,
                                retryOptions, "network failure while reading response", cancellationToken,
                                onRetry).ConfigureAwait(false);
                            continue;
                        } catch (IOException) when (CanRetry(requestSafety) && attempt < retryBudget) {
                            response.Dispose();
                            await DelayAfterTransportFailureAsync(method, uri, attempt, retryBudget,
                                retryOptions, "network failure while reading response", cancellationToken,
                                onRetry).ConfigureAwait(false);
                            continue;
                        } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested && CanRetry(requestSafety) && attempt < retryBudget) {
                            response.Dispose();
                            await DelayAfterTransportFailureAsync(method, uri, attempt, retryBudget,
                                retryOptions, "request timeout while reading response", cancellationToken,
                                onRetry).ConfigureAwait(false);
                            continue;
                        }
                    }

                    var (statusDelay, statusDelayStrategy) = GetRetryDelay(
                        response.Headers.RetryAfter, attempt, retryOptions);
                    onRetry?.Invoke(new GoogleWorkspaceRetryEvent(
                        method,
                        uri,
                        attempt + 1,
                        retryBudget,
                        $"HTTP {(int)response.StatusCode}",
                        statusDelay,
                        statusDelayStrategy));
                    response.Dispose();
                    await Task.Delay(statusDelay, cancellationToken)
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
            string trigger,
            CancellationToken cancellationToken,
            Action<GoogleWorkspaceRetryEvent>? onRetry) {
            var (delay, delayStrategy) = GetRetryDelay(null, attempt, retryOptions);
            onRetry?.Invoke(new GoogleWorkspaceRetryEvent(
                method,
                uri,
                attempt + 1,
                retryBudget,
                trigger,
                delay,
                delayStrategy));
            await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
        }

        private static bool CanRetry(GoogleWorkspaceRequestSafety requestSafety) {
            return requestSafety == GoogleWorkspaceRequestSafety.Safe
                || requestSafety == GoogleWorkspaceRequestSafety.Idempotent;
        }

        // Retry only the status codes Google APIs commonly use for throttling or transient infrastructure failures.
        private static bool ShouldRetry(HttpStatusCode statusCode) {
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
