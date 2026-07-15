namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Structured diagnostic entry that callers can forward to their own logging pipeline.
    /// </summary>
    public sealed class GoogleWorkspaceDiagnosticEntry {
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

        public string Code { get; }
        public TranslationSeverity Severity { get; }
        public string Feature { get; }
        public string Message { get; }
        public string Path { get; }
        public GoogleWorkspaceFailureKind? FailureKind { get; }
        public TranslationAction Action { get; }
        public int Count { get; }
        public string? TargetId { get; }
    }

    /// <summary>
    /// Helpers that translate reports and export exceptions into structured diagnostic entries.
    /// </summary>
    public static class GoogleWorkspaceDiagnosticsExtensions {
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

    public static class GoogleWorkspaceDiagnosticsDispatcher {
        public static void Emit(
            GoogleWorkspaceSessionOptions? sessionOptions,
            GoogleWorkspaceDiagnosticEntry entry) {
            if (entry == null) throw new ArgumentNullException(nameof(entry));
            sessionOptions?.DiagnosticSink?.Invoke(entry);
        }

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
