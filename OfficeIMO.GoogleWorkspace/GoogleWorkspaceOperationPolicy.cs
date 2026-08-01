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

    /// <summary>Semantic mutation kind supplied by the adapter independently of the HTTP verb.</summary>
    public enum GoogleWorkspaceMutationKind {
        Unspecified = 0,
        Create = 1,
        Update = 2,
        Delete = 3,
        Action = 4,
    }

    /// <summary>How a Google mutation's expected revision is actually enforced.</summary>
    public enum GoogleWorkspaceRevisionPreconditionKind {
        Unspecified = 0,
        ResourceAbsentCreate = 1,
        HttpEntityTag = 2,
        PayloadRevision = 3,
        Unavailable = 4,
        ResumableSessionState = 5,
    }

    /// <summary>Adapter-declared revision enforcement for one Google mutation.</summary>
    public sealed class GoogleWorkspaceRevisionPrecondition {
        private GoogleWorkspaceRevisionPrecondition(
            GoogleWorkspaceRevisionPreconditionKind kind,
            string? adapterExpectedRevision = null) {
            Kind = kind;
            AdapterExpectedRevision = adapterExpectedRevision;
        }

        /// <summary>Applies the policy's strong entity tag as an HTTP If-Match header.</summary>
        public static GoogleWorkspaceRevisionPrecondition HttpEntityTag { get; } =
            new GoogleWorkspaceRevisionPrecondition(GoogleWorkspaceRevisionPreconditionKind.HttpEntityTag);

        /// <summary>Declares that this API operation exposes no usable conditional revision precondition.</summary>
        public static GoogleWorkspaceRevisionPrecondition Unavailable { get; } =
            new GoogleWorkspaceRevisionPrecondition(GoogleWorkspaceRevisionPreconditionKind.Unavailable);

        /// <summary>Declares the exact revision already embedded by the adapter in the request payload.</summary>
        public static GoogleWorkspaceRevisionPrecondition PayloadRevision(string expectedRevision) {
            if (string.IsNullOrWhiteSpace(expectedRevision)) {
                throw new ArgumentException("The payload revision is required.", nameof(expectedRevision));
            }
            return new GoogleWorkspaceRevisionPrecondition(
                GoogleWorkspaceRevisionPreconditionKind.PayloadRevision,
                expectedRevision);
        }

        /// <summary>Declares the exact resumable-session state enforced by the request's content range.</summary>
        public static GoogleWorkspaceRevisionPrecondition ResumableSessionState(string expectedState) {
            if (string.IsNullOrWhiteSpace(expectedState)) {
                throw new ArgumentException("The resumable-session state is required.", nameof(expectedState));
            }
            return new GoogleWorkspaceRevisionPrecondition(
                GoogleWorkspaceRevisionPreconditionKind.ResumableSessionState,
                expectedState);
        }

        internal static GoogleWorkspaceRevisionPrecondition ResourceAbsentCreate { get; } =
            new GoogleWorkspaceRevisionPrecondition(GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate);

        public GoogleWorkspaceRevisionPreconditionKind Kind { get; }

        /// <summary>The exact revision or resumable-session state already carried by the adapter request.</summary>
        public string? AdapterExpectedRevision { get; }
    }

    /// <summary>Request facts supplied to the caller-owned policy provider before any mutation is sent.</summary>
    public sealed class GoogleWorkspaceOperationContext {
        internal GoogleWorkspaceOperationContext(string service, string method, string target,
            GoogleWorkspaceRequestSafety requestSafety, GoogleWorkspaceMutationKind mutationKind,
            GoogleWorkspaceRevisionPrecondition revisionPrecondition,
            bool potentialDataLoss, string? requestId) {
            Service = service; Method = method; Target = target; RequestSafety = requestSafety;
            MutationKind = mutationKind; RevisionPreconditionKind = revisionPrecondition.Kind;
            AdapterExpectedRevision = revisionPrecondition.AdapterExpectedRevision;
            PotentialDataLoss = potentialDataLoss; RequestId = requestId;
        }
        public string Service { get; }
        public string Method { get; }
        public string Target { get; }
        public GoogleWorkspaceRequestSafety RequestSafety { get; }
        /// <summary>Adapter-declared create, update, delete, or action semantics; never inferred from POST alone.</summary>
        public GoogleWorkspaceMutationKind MutationKind { get; }
        /// <summary>How the adapter declares that the policy revision will be enforced.</summary>
        public GoogleWorkspaceRevisionPreconditionKind RevisionPreconditionKind { get; }
        /// <summary>The exact revision already embedded in the request payload, when applicable.</summary>
        public string? AdapterExpectedRevision { get; }
        /// <summary>True when the request deletes remote data or otherwise has a transport-known loss risk.</summary>
        public bool PotentialDataLoss { get; }
        public string? RequestId { get; }
    }

    /// <summary>Explicit account, scope, revision, retry, rate, and loss contract for one cloud mutation.</summary>
    public sealed class GoogleWorkspaceOperationPolicy {
        /// <summary>Revision decision used for create operations whose target resource must not already exist.</summary>
        public const string ResourceAbsentForCreateRevision = "resource-absent-for-create";

        private const string ExplicitlyUnversionedPrefix = "explicitly-unversioned:";

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

        /// <summary>
        /// Records that the caller deliberately accepted a mutation for which the Google API exposes no usable
        /// conditional revision precondition. The named reason is retained in the operation receipt.
        /// </summary>
        public static string ExplicitlyUnversionedRevision(string reason) {
            if (string.IsNullOrWhiteSpace(reason)) {
                throw new ArgumentException("An explicitly unversioned mutation must name why no revision precondition is available.", nameof(reason));
            }
            return ExplicitlyUnversionedPrefix + reason.Trim();
        }

        internal static bool IsExplicitlyUnversioned(string expectedRevision) =>
            expectedRevision.StartsWith(ExplicitlyUnversionedPrefix, StringComparison.Ordinal);
    }

    /// <summary>Non-secret, caller-observable evidence for one attempted cloud mutation.</summary>
    public sealed class GoogleWorkspaceOperationReceipt {
        public GoogleWorkspaceOperationReceipt(GoogleWorkspaceOperationPolicy policy, string service,
            string method, string target, string? requestId, int retryCount, bool succeeded, string outcome) {
            Policy = policy; Service = service; Method = method; Target = target; RequestId = requestId;
            RetryCount = retryCount; Succeeded = succeeded; Outcome = outcome; CompletedAt = DateTimeOffset.UtcNow;
        }

        internal GoogleWorkspaceOperationReceipt(GoogleWorkspaceOperationPolicy policy, string service,
            string method, string target, string? requestId, int retryCount, bool succeeded, string outcome,
            GoogleWorkspaceMutationKind mutationKind,
            GoogleWorkspaceRevisionPreconditionKind revisionPreconditionKind, string? enforcedRevision) {
            Policy = policy; Service = service; Method = method; Target = target; RequestId = requestId;
            RetryCount = retryCount; Succeeded = succeeded; Outcome = outcome;
            MutationKind = mutationKind;
            RevisionPreconditionKind = revisionPreconditionKind; EnforcedRevision = enforcedRevision;
            CompletedAt = DateTimeOffset.UtcNow;
        }
        public GoogleWorkspaceOperationPolicy Policy { get; }
        public string Service { get; }
        public string Method { get; }
        public string Target { get; }
        public string? RequestId { get; }
        public int RetryCount { get; }
        public bool Succeeded { get; }
        public string Outcome { get; }
        /// <summary>The adapter-declared semantic mutation represented by this receipt.</summary>
        public GoogleWorkspaceMutationKind MutationKind { get; }
        /// <summary>How the expected revision was enforced for this attempted mutation.</summary>
        public GoogleWorkspaceRevisionPreconditionKind RevisionPreconditionKind { get; }
        /// <summary>The revision actually enforced by HTTP or payload precondition, when one was available.</summary>
        public string? EnforcedRevision { get; }
        public DateTimeOffset CompletedAt { get; }
    }

    /// <summary>
    /// Indicates that a remote mutation outcome is known, but its caller-provided receipt sink failed.
    /// </summary>
    [Serializable]
    public sealed class GoogleWorkspaceReceiptPersistenceException : Exception {
        /// <summary>Key used when a receipt failure is attached to the original remote-operation exception.</summary>
        public const string ExceptionDataKey = "OfficeIMO.GoogleWorkspace.ReceiptPersistenceFailure";

        internal GoogleWorkspaceReceiptPersistenceException(
            GoogleWorkspaceOperationReceipt receipt,
            bool remoteOperationSucceeded,
            Exception innerException)
            : base(remoteOperationSucceeded
                ? "The Google Workspace mutation succeeded remotely, but its operation receipt could not be persisted. Do not retry the mutation without reconciling the remote resource."
                : "The Google Workspace mutation failed and its operation receipt could not be persisted.", innerException) {
            Receipt = receipt;
            RemoteOperationSucceeded = remoteOperationSucceeded;
        }

        /// <summary>The receipt that the caller-provided sink failed to persist.</summary>
        public GoogleWorkspaceOperationReceipt Receipt { get; }

        /// <summary>True when the remote mutation completed successfully before receipt persistence failed.</summary>
        public bool RemoteOperationSucceeded { get; }
    }
}
