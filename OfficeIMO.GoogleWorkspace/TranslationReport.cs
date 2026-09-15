namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Action selected by a translator for a source feature.
    /// </summary>
    public enum TranslationAction {
        /// <summary>No target action has been selected.</summary>
        None = 0,
        /// <summary>Preserve the source feature in its native target form.</summary>
        Preserve = 1,
        /// <summary>Omit the unsupported source feature.</summary>
        Skip = 2,
        /// <summary>Convert the feature to a simpler editable representation.</summary>
        Flatten = 3,
        /// <summary>Render the feature as an image.</summary>
        Rasterize = 4,
        /// <summary>Reject the translation before mutation.</summary>
        Fail = 5,
    }

    /// <summary>
    /// Standard fidelity report shared across exporter packages.
    /// </summary>
    public sealed class TranslationReport {
        private readonly List<TranslationNotice> _notices;
        private readonly IReadOnlyList<TranslationNotice> _readOnlyNotices;
        private readonly bool _isReadOnly;

        /// <summary>Creates an empty, mutable translation report.</summary>
        public TranslationReport() : this(Array.Empty<TranslationNotice>(), false) { }

        private TranslationReport(IEnumerable<TranslationNotice> notices, bool isReadOnly) {
            _notices = notices.ToList();
            _readOnlyNotices = _notices.AsReadOnly();
            _isReadOnly = isReadOnly;
        }

        /// <summary>Gets the notices in the order in which translators recorded them.</summary>
        public IReadOnlyList<TranslationNotice> Notices => _readOnlyNotices;
        /// <summary>Gets whether the report contains at least one warning or error.</summary>
        public bool HasWarnings => _notices.Any(n => n.Severity >= TranslationSeverity.Warning);
        /// <summary>Gets whether the report contains at least one error.</summary>
        public bool HasErrors => _notices.Any(n => n.Severity >= TranslationSeverity.Error);

        /// <summary>Adds a structured translation notice.</summary>
        /// <param name="severity">Impact of the notice.</param>
        /// <param name="feature">Source feature to which the notice applies.</param>
        /// <param name="message">Human-readable explanation.</param>
        /// <param name="path">Optional source-object path.</param>
        /// <param name="code">Stable diagnostic code, or <see langword="null"/> to derive one from <paramref name="feature"/>.</param>
        /// <param name="action">Action selected for the feature.</param>
        /// <param name="count">Number of equivalent occurrences represented by the notice; values below one become one.</param>
        /// <param name="targetId">Optional remote target identifier.</param>
        public void Add(
            TranslationSeverity severity,
            string feature,
            string message,
            string path = "",
            string? code = null,
            TranslationAction action = TranslationAction.None,
            int count = 1,
            string? targetId = null) {
            EnsureMutable();
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

        /// <summary>Adds a notice unless the report already contains the same severity, code, feature, message, and path.</summary>
        /// <param name="severity">Impact of the notice.</param>
        /// <param name="feature">Source feature to which the notice applies.</param>
        /// <param name="message">Human-readable explanation.</param>
        /// <param name="path">Optional source-object path.</param>
        /// <param name="code">Stable diagnostic code, or <see langword="null"/> to derive one from <paramref name="feature"/>.</param>
        /// <param name="action">Action selected for the feature.</param>
        /// <param name="count">Number of equivalent occurrences represented by the notice.</param>
        /// <param name="targetId">Optional remote target identifier.</param>
        public void AddUnique(
            TranslationSeverity severity,
            string feature,
            string message,
            string path = "",
            string? code = null,
            TranslationAction action = TranslationAction.None,
            int count = 1,
            string? targetId = null) {
            EnsureMutable();
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

        /// <summary>Creates an independent read-only snapshot of the current notices.</summary>
        public TranslationReport CreateReadOnlySnapshot() => new TranslationReport(_notices, true);

        private void EnsureMutable() {
            if (_isReadOnly) throw new InvalidOperationException("This translation report is a read-only planning snapshot.");
        }
    }

    /// <summary>
    /// A single fidelity or planning notice.
    /// </summary>
    public sealed class TranslationNotice {
        /// <summary>Creates an immutable translation notice.</summary>
        /// <param name="path">Source-object path, or an empty string when not applicable.</param>
        /// <param name="feature">Source feature described by the notice.</param>
        /// <param name="severity">Impact of the notice.</param>
        /// <param name="message">Human-readable explanation.</param>
        /// <param name="code">Stable machine-readable diagnostic code.</param>
        /// <param name="action">Action selected for the feature.</param>
        /// <param name="count">Number of occurrences represented by the notice.</param>
        /// <param name="targetId">Optional remote target identifier.</param>
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

        /// <summary>Gets the stable machine-readable diagnostic code.</summary>
        public string Code { get; }
        /// <summary>Gets the source-object path, or an empty string when not applicable.</summary>
        public string Path { get; }
        /// <summary>Gets the source feature described by the notice.</summary>
        public string Feature { get; }
        /// <summary>Gets the impact of the notice.</summary>
        public TranslationSeverity Severity { get; }
        /// <summary>Gets the human-readable explanation.</summary>
        public string Message { get; }
        /// <summary>Gets the target action selected for the feature.</summary>
        public TranslationAction Action { get; }
        /// <summary>Gets the number of equivalent occurrences represented by the notice.</summary>
        public int Count { get; }
        /// <summary>Gets the related remote target identifier, when available.</summary>
        public string? TargetId { get; }
    }

    /// <summary>
    /// Severity levels for translation notices.
    /// </summary>
    public enum TranslationSeverity {
        /// <summary>Informational notice that does not indicate fidelity loss.</summary>
        Info = 0,
        /// <summary>Notice about potential fidelity loss or an operator decision.</summary>
        Warning = 1,
        /// <summary>Notice describing a condition that normally blocks the operation.</summary>
        Error = 2,
    }
}
