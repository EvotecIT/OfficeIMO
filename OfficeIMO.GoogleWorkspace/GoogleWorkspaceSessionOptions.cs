using System.Net.Http;

namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Session-level options shared by Google Workspace exporters.
    /// </summary>
    public sealed class GoogleWorkspaceSessionOptions {
        public string ApplicationName { get; set; } = "OfficeIMO";
        public string? DefaultDriveId { get; set; }
        public string? DefaultFolderId { get; set; }
        public string? SubjectUser { get; set; }
        public bool UseDomainWideDelegation { get; set; }
        public HttpClient? HttpClient { get; set; }
        public TimeSpan RequestTimeout { get; set; } = TimeSpan.FromSeconds(100);
        public int MaxRetryCount { get; set; } = 3;
        public TimeSpan RetryBaseDelay { get; set; } = TimeSpan.FromMilliseconds(200);
        public TimeSpan RetryMaxDelay { get; set; } = TimeSpan.FromSeconds(5);
        public Action<GoogleWorkspaceDiagnosticEntry>? DiagnosticSink { get; set; }
    }
}
