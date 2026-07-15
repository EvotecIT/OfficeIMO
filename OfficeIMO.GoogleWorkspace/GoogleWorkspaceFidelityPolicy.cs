namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Requested behavior for a source feature without a native target equivalent.
    /// </summary>
    public enum UnsupportedFeatureMode {
        Error = 0,
        WarnAndSkip = 1,
        Flatten = 2,
        Rasterize = 3,
    }

    /// <summary>
    /// Determines which diagnostics stop an operation before any Google mutation.
    /// </summary>
    public enum GoogleWorkspacePreflightMode {
        ReportOnly = 0,
        FailOnErrors = 1,
        FailOnWarnings = 2,
    }

    /// <summary>
    /// Shared preflight policy used by domain translators.
    /// </summary>
    public sealed class GoogleWorkspaceFidelityPolicy {
        public GoogleWorkspacePreflightMode PreflightMode { get; set; } = GoogleWorkspacePreflightMode.FailOnErrors;
        public ISet<string> AcceptedDiagnosticCodes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    public static class GoogleWorkspacePreflight {
        public static void Validate(TranslationReport report, GoogleWorkspaceFidelityPolicy policy) {
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (policy == null) throw new ArgumentNullException(nameof(policy));

            if (policy.PreflightMode == GoogleWorkspacePreflightMode.ReportOnly) {
                return;
            }

            var blocking = report.Notices
                .Where(notice => !policy.AcceptedDiagnosticCodes.Contains(notice.Code))
                .Where(notice => policy.PreflightMode == GoogleWorkspacePreflightMode.FailOnWarnings
                    ? notice.Severity >= TranslationSeverity.Warning
                    : notice.Severity >= TranslationSeverity.Error)
                .ToArray();
            if (blocking.Length == 0) {
                return;
            }

            string summary = string.Join(", ", blocking.Select(notice => notice.Code).Distinct(StringComparer.Ordinal));
            throw new GoogleWorkspacePreflightException(
                $"Google Workspace preflight blocked the operation before mutation because of: {summary}.",
                report,
                blocking);
        }
    }

    public sealed class GoogleWorkspacePreflightException : InvalidOperationException {
        public GoogleWorkspacePreflightException(
            string message,
            TranslationReport report,
            IReadOnlyList<TranslationNotice> blockingNotices)
            : base(message) {
            Report = report ?? throw new ArgumentNullException(nameof(report));
            BlockingNotices = blockingNotices ?? throw new ArgumentNullException(nameof(blockingNotices));
        }

        public TranslationReport Report { get; }
        public IReadOnlyList<TranslationNotice> BlockingNotices { get; }
    }
}
