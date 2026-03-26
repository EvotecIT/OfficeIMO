namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Standard fidelity report shared across exporter packages.
    /// </summary>
    public sealed class TranslationReport {
        private readonly List<TranslationNotice> _notices = new List<TranslationNotice>();

        public IReadOnlyList<TranslationNotice> Notices => _notices;
        public bool HasWarnings => _notices.Any(n => n.Severity >= TranslationSeverity.Warning);
        public bool HasErrors => _notices.Any(n => n.Severity >= TranslationSeverity.Error);

        public void Add(TranslationSeverity severity, string feature, string message, string path = "") {
            _notices.Add(new TranslationNotice(path, feature, severity, message));
        }

        public void AddUnique(TranslationSeverity severity, string feature, string message, string path = "") {
            if (_notices.Any(n =>
                n.Severity == severity
                && string.Equals(n.Feature, feature, StringComparison.Ordinal)
                && string.Equals(n.Message, message, StringComparison.Ordinal)
                && string.Equals(n.Path, path, StringComparison.Ordinal))) {
                return;
            }

            _notices.Add(new TranslationNotice(path, feature, severity, message));
        }
    }

    /// <summary>
    /// A single fidelity or planning notice.
    /// </summary>
    public sealed class TranslationNotice {
        public TranslationNotice(string path, string feature, TranslationSeverity severity, string message) {
            Path = path ?? string.Empty;
            Feature = feature ?? string.Empty;
            Severity = severity;
            Message = message ?? string.Empty;
        }

        public string Path { get; }
        public string Feature { get; }
        public TranslationSeverity Severity { get; }
        public string Message { get; }
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
