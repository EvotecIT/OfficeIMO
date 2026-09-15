using System.Net.Http;

namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Session-level options shared by Google Workspace exporters.
    /// </summary>
    public sealed class GoogleWorkspaceSessionOptions {
        /// <summary>Gets or sets the application name sent in Google API requests. The default is <c>OfficeIMO</c>.</summary>
        public string ApplicationName { get; set; } = "OfficeIMO";
        /// <summary>Gets or sets the stable end-user identifier sent for provider-side quota accounting.</summary>
        public string? QuotaUser { get; set; }
        /// <summary>Gets or sets the requested Google Cloud consumer project sent for provider-side quota and billing evaluation.</summary>
        public string? QuotaProject { get; set; }
        /// <summary>Gets or sets a factory for correlation identifiers attached to outgoing operations.</summary>
        public Func<string>? RequestIdFactory { get; set; }
        /// <summary>Gets or sets the default shared-drive identifier used when an operation omits one.</summary>
        public string? DefaultDriveId { get; set; }
        /// <summary>Gets or sets the default parent-folder identifier used when an operation omits one.</summary>
        public string? DefaultFolderId { get; set; }
        /// <summary>Gets or sets the Workspace user requested for service-account impersonation.</summary>
        /// <remarks>The user is included as the JWT subject only when <see cref="UseDomainWideDelegation"/> is enabled; successful impersonation also requires provider authorization.</remarks>
        public string? SubjectUser { get; set; }
        /// <summary>Gets or sets whether service-account credentials request domain-wide delegation.</summary>
        public bool UseDomainWideDelegation { get; set; }
        /// <summary>Gets or sets the HTTP client used for Google API requests. When absent, the transport owns a private client.</summary>
        public HttpClient? HttpClient { get; set; }
        /// <summary>Gets or sets the timeout applied to each request attempt. The default is 100 seconds.</summary>
        public TimeSpan RequestTimeout { get; set; } = TimeSpan.FromSeconds(100);
        /// <summary>Gets or sets the maximum number of retries after the initial request. The default is 3.</summary>
        public int MaxRetryCount { get; set; } = 3;
        /// <summary>Gets or sets the initial delay for exponential retry backoff. The default is 200 milliseconds.</summary>
        public TimeSpan RetryBaseDelay { get; set; } = TimeSpan.FromMilliseconds(200);
        /// <summary>Gets or sets the upper bound for one retry delay. The default is 5 seconds.</summary>
        public TimeSpan RetryMaxDelay { get; set; } = TimeSpan.FromSeconds(5);
        /// <summary>Gets or sets the overall elapsed-time budget for one operation, including the initial attempt, retries, delays, and response processing. The default is 2 minutes.</summary>
        public TimeSpan MaxRetryElapsedTime { get; set; } = TimeSpan.FromMinutes(2);
        /// <summary>Gets or sets how server rate-limit guidance influences retry delays.</summary>
        public GoogleWorkspaceRateLimitPolicy RateLimitPolicy { get; set; } = GoogleWorkspaceRateLimitPolicy.HonorRetryAfter;
        /// <summary>Gets or sets the required provider-verified account identity for acquired credentials.</summary>
        public string? ExpectedAccount { get; set; }
        /// <summary>Gets or sets a callback that supplies safety, retry, and precondition policy for an operation.</summary>
        public Func<GoogleWorkspaceOperationContext, GoogleWorkspaceOperationPolicy>? OperationPolicyProvider { get; set; }
        /// <summary>Gets or sets a sink that receives the final receipt for each mutation operation.</summary>
        public Action<GoogleWorkspaceOperationReceipt>? OperationReceiptSink { get; set; }
        /// <summary>Gets or sets a sink for structured diagnostics emitted during Workspace operations.</summary>
        public Action<GoogleWorkspaceDiagnosticEntry>? DiagnosticSink { get; set; }
    }
}
