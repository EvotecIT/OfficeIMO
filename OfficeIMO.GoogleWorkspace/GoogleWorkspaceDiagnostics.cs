namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Structured diagnostic entry that callers can forward to their own logging pipeline.
    /// </summary>
    public sealed class GoogleWorkspaceDiagnosticEntry {
        /// <summary>Creates a normalized structured diagnostic entry.</summary>
        /// <param name="severity">Impact of the diagnostic.</param>
        /// <param name="feature">Feature that emitted the diagnostic.</param>
        /// <param name="message">Human-readable explanation.</param>
        /// <param name="path">Optional source-object path.</param>
        /// <param name="failureKind">Related export-failure category, when applicable.</param>
        /// <param name="code">Stable diagnostic code, or <see langword="null"/> to derive one from <paramref name="feature"/>.</param>
        /// <param name="action">Target action selected for the feature.</param>
        /// <param name="count">Number of equivalent occurrences represented by the entry.</param>
        /// <param name="targetId">Optional remote target identifier.</param>
        public GoogleWorkspaceDiagnosticEntry(
            TranslationSeverity severity,
            string feature,
            string message,
            string path = "",
            GoogleWorkspaceFailureKind? failureKind = null,
            string? code = null,
            TranslationAction action = TranslationAction.None,
            int count = 1,
            string? targetId = null) {
            Severity = severity;
            Feature = feature ?? string.Empty;
            Message = message ?? string.Empty;
            Path = path ?? string.Empty;
            FailureKind = failureKind;
            Code = GoogleWorkspaceDiagnosticCodes.Resolve(code, feature);
            Action = action;
            Count = Math.Max(1, count);
            TargetId = targetId;
        }

        /// <summary>Gets the stable machine-readable diagnostic code.</summary>
        public string Code { get; }
        /// <summary>Gets the impact of the diagnostic.</summary>
        public TranslationSeverity Severity { get; }
        /// <summary>Gets the feature that emitted the diagnostic.</summary>
        public string Feature { get; }
        /// <summary>Gets the human-readable explanation.</summary>
        public string Message { get; }
        /// <summary>Gets the source-object path, or an empty string when not applicable.</summary>
        public string Path { get; }
        /// <summary>Gets the related export-failure category, when applicable.</summary>
        public GoogleWorkspaceFailureKind? FailureKind { get; }
        /// <summary>Gets the target action selected for the feature.</summary>
        public TranslationAction Action { get; }
        /// <summary>Gets the number of equivalent occurrences represented by the entry.</summary>
        public int Count { get; }
        /// <summary>Gets the related remote target identifier, when available.</summary>
        public string? TargetId { get; }
    }

    /// <summary>
    /// Helpers that translate reports and export exceptions into structured diagnostic entries.
    /// </summary>
    public static class GoogleWorkspaceDiagnosticsExtensions {
        /// <summary>Converts every notice in a translation report to a structured diagnostic entry.</summary>
        /// <param name="report">Report to convert.</param>
        /// <returns>A snapshot preserving notice order and metadata.</returns>
        public static IReadOnlyList<GoogleWorkspaceDiagnosticEntry> ToDiagnosticEntries(this TranslationReport report) {
            if (report == null) throw new ArgumentNullException(nameof(report));

            return report.Notices
                .Select(notice => new GoogleWorkspaceDiagnosticEntry(
                    notice.Severity,
                    notice.Feature,
                    notice.Message,
                    notice.Path,
                    code: notice.Code,
                    action: notice.Action,
                    count: notice.Count,
                    targetId: notice.TargetId))
                .ToArray();
        }

        /// <summary>Converts an export exception and its report to structured diagnostics.</summary>
        /// <param name="exception">Export failure to convert.</param>
        /// <returns>A leading export-failure entry followed by the report notices.</returns>
        public static IReadOnlyList<GoogleWorkspaceDiagnosticEntry> ToDiagnosticEntries(this GoogleWorkspaceExportException exception) {
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            var entries = new List<GoogleWorkspaceDiagnosticEntry> {
                new GoogleWorkspaceDiagnosticEntry(
                    TranslationSeverity.Error,
                    "ExportFailure",
                    exception.Message,
                    failureKind: exception.FailureKind,
                    code: "WORKSPACE.EXPORT.FAILED",
                    action: TranslationAction.Fail)
            };

            entries.AddRange(exception.Report.Notices.Select(notice => new GoogleWorkspaceDiagnosticEntry(
                notice.Severity,
                notice.Feature,
                notice.Message,
                notice.Path,
                exception.FailureKind,
                notice.Code,
                notice.Action,
                notice.Count,
                notice.TargetId)));

            return entries;
        }
    }

    /// <summary>Records translation notices and forwards matching structured entries to a session sink.</summary>
    public static class GoogleWorkspaceDiagnosticsDispatcher {
        /// <summary>Sends one entry to the configured diagnostic sink, when present.</summary>
        /// <param name="sessionOptions">Session containing the optional sink.</param>
        /// <param name="entry">Entry to emit.</param>
        public static void Emit(
            GoogleWorkspaceSessionOptions? sessionOptions,
            GoogleWorkspaceDiagnosticEntry entry) {
            if (entry == null) throw new ArgumentNullException(nameof(entry));
            sessionOptions?.DiagnosticSink?.Invoke(entry);
        }

        /// <summary>Adds a notice to a report and emits the equivalent structured diagnostic.</summary>
        public static void Add(
            TranslationReport report,
            GoogleWorkspaceSessionOptions? sessionOptions,
            TranslationSeverity severity,
            string feature,
            string message,
            string path = "",
            GoogleWorkspaceFailureKind? failureKind = null,
            string? code = null,
            TranslationAction action = TranslationAction.None,
            int count = 1,
            string? targetId = null) {
            if (report == null) throw new ArgumentNullException(nameof(report));

            report.Add(severity, feature, message, path, code, action, count, targetId);
            Emit(sessionOptions, new GoogleWorkspaceDiagnosticEntry(severity, feature, message, path, failureKind, code, action, count, targetId));
        }

        /// <summary>Adds and emits a diagnostic unless the same severity, feature, message, and path already exist.</summary>
        public static void AddUnique(
            TranslationReport report,
            GoogleWorkspaceSessionOptions? sessionOptions,
            TranslationSeverity severity,
            string feature,
            string message,
            string path = "",
            GoogleWorkspaceFailureKind? failureKind = null,
            string? code = null,
            TranslationAction action = TranslationAction.None,
            int count = 1,
            string? targetId = null) {
            if (report == null) throw new ArgumentNullException(nameof(report));

            if (report.Notices.Any(n =>
                n.Severity == severity
                && string.Equals(n.Feature, feature, StringComparison.Ordinal)
                && string.Equals(n.Message, message, StringComparison.Ordinal)
                && string.Equals(n.Path, path, StringComparison.Ordinal))) {
                return;
            }

            report.AddUnique(severity, feature, message, path, code, action, count, targetId);
            Emit(sessionOptions, new GoogleWorkspaceDiagnosticEntry(severity, feature, message, path, failureKind, code, action, count, targetId));
        }
    }
}
