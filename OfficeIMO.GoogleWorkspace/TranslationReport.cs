namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Action selected by a translator for a source feature.
    /// </summary>
    public enum TranslationAction {
        None = 0,
        Preserve = 1,
        Skip = 2,
        Flatten = 3,
        Rasterize = 4,
        Fail = 5,
    }

    /// <summary>
    /// Standard fidelity report shared across exporter packages.
    /// </summary>
    public sealed class TranslationReport {
        private readonly List<TranslationNotice> _notices = new List<TranslationNotice>();

        public IReadOnlyList<TranslationNotice> Notices => _notices;
        public bool HasWarnings => _notices.Any(n => n.Severity >= TranslationSeverity.Warning);
        public bool HasErrors => _notices.Any(n => n.Severity >= TranslationSeverity.Error);

        public void Add(
            TranslationSeverity severity,
            string feature,
            string message,
            string path = "",
            string? code = null,
            TranslationAction action = TranslationAction.None,
            int count = 1,
            string? targetId = null) {
            _notices.Add(new TranslationNotice(
                path,
                feature,
                severity,
                message,
                GoogleWorkspaceDiagnosticCodes.Resolve(code, feature),
                action,
                count,
                targetId));
        }

        public void AddUnique(
            TranslationSeverity severity,
            string feature,
            string message,
            string path = "",
            string? code = null,
            TranslationAction action = TranslationAction.None,
            int count = 1,
            string? targetId = null) {
            string resolvedCode = GoogleWorkspaceDiagnosticCodes.Resolve(code, feature);
            if (_notices.Any(n =>
                n.Severity == severity
                && string.Equals(n.Code, resolvedCode, StringComparison.Ordinal)
                && string.Equals(n.Feature, feature, StringComparison.Ordinal)
                && string.Equals(n.Message, message, StringComparison.Ordinal)
                && string.Equals(n.Path, path, StringComparison.Ordinal))) {
                return;
            }

            _notices.Add(new TranslationNotice(path, feature, severity, message, resolvedCode, action, count, targetId));
        }
    }

    /// <summary>
    /// A single fidelity or planning notice.
    /// </summary>
    public sealed class TranslationNotice {
        public TranslationNotice(
            string path,
            string feature,
            TranslationSeverity severity,
            string message,
            string code,
            TranslationAction action,
            int count,
            string? targetId) {
            Path = path ?? string.Empty;
            Feature = feature ?? string.Empty;
            Severity = severity;
            Message = message ?? string.Empty;
            Code = code ?? string.Empty;
            Action = action;
            Count = Math.Max(1, count);
            TargetId = targetId;
        }

        public string Code { get; }
        public string Path { get; }
        public string Feature { get; }
        public TranslationSeverity Severity { get; }
        public string Message { get; }
        public TranslationAction Action { get; }
        public int Count { get; }
        public string? TargetId { get; }
    }

    /// <summary>
    /// Severity levels for translation notices.
    /// </summary>
    public enum TranslationSeverity {
        Info = 0,
        Warning = 1,
        Error = 2,
    }
}
