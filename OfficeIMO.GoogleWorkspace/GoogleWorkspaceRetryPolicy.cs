using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;

namespace OfficeIMO.GoogleWorkspace {
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
        public static async Task<HttpResponseMessage> SendAsync(
            HttpClient client,
            Func<HttpRequestMessage> requestFactory,
            GoogleWorkspaceRetryOptions retryOptions,
            CancellationToken cancellationToken,
            Action<GoogleWorkspaceRetryEvent>? onRetry = null) {
            if (retryOptions == null) throw new ArgumentNullException(nameof(retryOptions));
            int retryBudget = retryOptions.MaxRetryCount;

            for (int attempt = 0; ; attempt++) {
                using (var request = requestFactory()) {
                    string method = request.Method.Method;
                    string uri = request.RequestUri?.AbsoluteUri ?? string.Empty;

                    try {
                        var response = await client.SendAsync(request, cancellationToken).ConfigureAwait(false);
                        if (!ShouldRetry(response.StatusCode) || attempt >= retryBudget) {
                            return response;
                        }

                        var (delay, delayStrategy) = GetRetryDelay(response.Headers.RetryAfter, attempt, retryOptions);
                        onRetry?.Invoke(new GoogleWorkspaceRetryEvent(
                            method,
                            uri,
                            attempt + 1,
                            retryBudget,
                            $"HTTP {(int)response.StatusCode}",
                            delay,
                            delayStrategy));
                        response.Dispose();
                        await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
                    } catch (HttpRequestException) when (attempt < retryBudget) {
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
                    } catch (TaskCanceledException) when (!cancellationToken.IsCancellationRequested && attempt < retryBudget) {
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
                    }
                }
            }
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
