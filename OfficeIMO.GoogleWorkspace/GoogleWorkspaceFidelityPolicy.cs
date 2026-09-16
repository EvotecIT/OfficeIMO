namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Requested behavior for a source feature without a native target equivalent.
    /// </summary>
    public enum UnsupportedFeatureMode {
        /// <summary>Report an error; the configured preflight policy determines whether it blocks mutation.</summary>
        Error = 0,
        /// <summary>Report a warning and omit the feature.</summary>
        WarnAndSkip = 1,
        /// <summary>Replace the feature with a simpler editable representation.</summary>
        Flatten = 2,
        /// <summary>Replace the feature with a rendered image.</summary>
        Rasterize = 3,
    }

    /// <summary>
    /// Determines which diagnostics stop an operation before any Google mutation.
    /// </summary>
    public enum GoogleWorkspacePreflightMode {
        /// <summary>Return diagnostics without blocking mutation.</summary>
        ReportOnly = 0,
        /// <summary>Block mutation when the report contains an unaccepted error.</summary>
        FailOnErrors = 1,
        /// <summary>Block mutation when the report contains an unaccepted warning or error.</summary>
        FailOnWarnings = 2,
    }

    /// <summary>
    /// Shared preflight policy used by domain translators.
    /// </summary>
    public sealed class GoogleWorkspaceFidelityPolicy {
        /// <summary>Gets or sets the severity threshold that blocks remote mutation. The default is <see cref="GoogleWorkspacePreflightMode.FailOnErrors"/>.</summary>
        public GoogleWorkspacePreflightMode PreflightMode { get; set; } = GoogleWorkspacePreflightMode.FailOnErrors;
        /// <summary>Gets diagnostic codes that callers explicitly accept and exclude from preflight blocking.</summary>
        public ISet<string> AcceptedDiagnosticCodes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Applies fidelity policy to a translation report before any remote mutation.</summary>
    public static class GoogleWorkspacePreflight {
        /// <summary>Checks a translation report and rejects notices that exceed the configured threshold.</summary>
        /// <param name="report">Report produced during translation planning.</param>
        /// <param name="policy">Policy controlling accepted codes and blocking severity.</param>
        /// <exception cref="GoogleWorkspacePreflightException">One or more unaccepted notices block the operation.</exception>
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

    /// <summary>Represents a preflight rejection before any Google Workspace mutation occurs.</summary>
    public sealed class GoogleWorkspacePreflightException : InvalidOperationException {
        /// <summary>Creates an exception for the notices that blocked preflight.</summary>
        /// <param name="message">Human-readable rejection summary.</param>
        /// <param name="report">Complete translation report.</param>
        /// <param name="blockingNotices">Unaccepted notices that crossed the policy threshold.</param>
        public GoogleWorkspacePreflightException(
            string message,
            TranslationReport report,
            IReadOnlyList<TranslationNotice> blockingNotices)
            : base(message) {
            Report = report ?? throw new ArgumentNullException(nameof(report));
            BlockingNotices = blockingNotices ?? throw new ArgumentNullException(nameof(blockingNotices));
        }

        /// <summary>Gets the complete translation report available at rejection time.</summary>
        public TranslationReport Report { get; }
        /// <summary>Gets the subset of notices that blocked the operation.</summary>
        public IReadOnlyList<TranslationNotice> BlockingNotices { get; }
    }
}
