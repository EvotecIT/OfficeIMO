namespace OfficeIMO.GoogleWorkspace {
    /// <summary>Caller decision for a mutation whose outcome can omit, replace, or delete data.</summary>
    public enum GoogleWorkspaceDataLossDecision {
        RejectPotentialLoss = 0,
        AcceptSpecifiedLoss = 1,
    }

    /// <summary>Caller-selected behavior when Google reports quota throttling.</summary>
    public enum GoogleWorkspaceRateLimitPolicy {
        HonorRetryAfter = 0,
        FailFast = 1,
    }

    /// <summary>Request facts supplied to the caller-owned policy provider before any mutation is sent.</summary>
    public sealed class GoogleWorkspaceOperationContext {
        internal GoogleWorkspaceOperationContext(string service, string method, string target,
            GoogleWorkspaceRequestSafety requestSafety, bool potentialDataLoss, string? requestId) {
            Service = service; Method = method; Target = target; RequestSafety = requestSafety;
            PotentialDataLoss = potentialDataLoss; RequestId = requestId;
        }
        public string Service { get; }
        public string Method { get; }
        public string Target { get; }
        public GoogleWorkspaceRequestSafety RequestSafety { get; }
        /// <summary>True when the request deletes remote data or otherwise has a transport-known loss risk.</summary>
        public bool PotentialDataLoss { get; }
        public string? RequestId { get; }
    }

    /// <summary>Explicit account, scope, revision, retry, rate, and loss contract for one cloud mutation.</summary>
    public sealed class GoogleWorkspaceOperationPolicy {
        public GoogleWorkspaceOperationPolicy(string account, IEnumerable<string> scopes, string target,
            string expectedRevision, int maxRetryCount, TimeSpan maxRetryElapsedTime,
            GoogleWorkspaceRateLimitPolicy rateLimitPolicy, GoogleWorkspaceDataLossDecision dataLossDecision,
            string? acceptedLoss = null) {
            if (string.IsNullOrWhiteSpace(account)) throw new ArgumentException("The expected Google account is required.", nameof(account));
            if (scopes == null) throw new ArgumentNullException(nameof(scopes));
            string[] materialized = scopes.Where(scope => !string.IsNullOrWhiteSpace(scope)).Distinct(StringComparer.Ordinal).ToArray();
            if (materialized.Length == 0) throw new ArgumentException("At least one OAuth scope is required.", nameof(scopes));
            if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("The mutation target is required.", nameof(target));
            if (string.IsNullOrWhiteSpace(expectedRevision)) throw new ArgumentException("An expected revision decision is required; use an explicit create or unversioned marker when applicable.", nameof(expectedRevision));
            if (maxRetryCount < 0) throw new ArgumentOutOfRangeException(nameof(maxRetryCount));
            if (maxRetryElapsedTime <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(maxRetryElapsedTime));
            if (dataLossDecision == GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss && string.IsNullOrWhiteSpace(acceptedLoss)) throw new ArgumentException("Accepted loss must be named explicitly.", nameof(acceptedLoss));
            Account = account; Scopes = materialized; Target = target; ExpectedRevision = expectedRevision;
            MaxRetryCount = maxRetryCount; MaxRetryElapsedTime = maxRetryElapsedTime;
            RateLimitPolicy = rateLimitPolicy; DataLossDecision = dataLossDecision; AcceptedLoss = acceptedLoss;
        }
        public string Account { get; }
        public IReadOnlyList<string> Scopes { get; }
        public string Target { get; }
        public string ExpectedRevision { get; }
        public int MaxRetryCount { get; }
        public TimeSpan MaxRetryElapsedTime { get; }
        public GoogleWorkspaceRateLimitPolicy RateLimitPolicy { get; }
        public GoogleWorkspaceDataLossDecision DataLossDecision { get; }
        public string? AcceptedLoss { get; }
    }

    /// <summary>Non-secret, caller-observable evidence for one attempted cloud mutation.</summary>
    public sealed class GoogleWorkspaceOperationReceipt {
        public GoogleWorkspaceOperationReceipt(GoogleWorkspaceOperationPolicy policy, string service,
            string method, string target, string? requestId, int retryCount, bool succeeded, string outcome) {
            Policy = policy; Service = service; Method = method; Target = target; RequestId = requestId;
            RetryCount = retryCount; Succeeded = succeeded; Outcome = outcome; CompletedAt = DateTimeOffset.UtcNow;
        }
        public GoogleWorkspaceOperationPolicy Policy { get; }
        public string Service { get; }
        public string Method { get; }
        public string Target { get; }
        public string? RequestId { get; }
        public int RetryCount { get; }
        public bool Succeeded { get; }
        public string Outcome { get; }
        public DateTimeOffset CompletedAt { get; }
    }
}
