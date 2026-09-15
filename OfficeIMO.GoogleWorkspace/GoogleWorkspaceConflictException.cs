namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Raised before mutation when a remote resource no longer matches the caller's observed version.
    /// </summary>
    public sealed class GoogleWorkspaceConflictException : InvalidOperationException {
        /// <summary>Initializes a conflict with the expected and observed versions of the target resource.</summary>
        /// <param name="message">Message describing the rejected mutation.</param>
        /// <param name="targetId">Identifier of the resource whose version changed.</param>
        /// <param name="expectedVersion">Version expected by the caller, or <see langword="null"/> when unavailable.</param>
        /// <param name="actualVersion">Version observed before mutation, or <see langword="null"/> when unavailable.</param>
        /// <param name="report">Translation report containing the associated conflict diagnostics.</param>
        public GoogleWorkspaceConflictException(
            string message,
            string targetId,
            string? expectedVersion,
            string? actualVersion,
            TranslationReport report)
            : base(message) {
            TargetId = targetId ?? throw new ArgumentNullException(nameof(targetId));
            ExpectedVersion = expectedVersion;
            ActualVersion = actualVersion;
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        /// <summary>Gets the identifier of the resource whose mutation was rejected.</summary>
        public string TargetId { get; }
        /// <summary>Gets the resource version expected by the caller, when known.</summary>
        public string? ExpectedVersion { get; }
        /// <summary>Gets the resource version observed before mutation, when known.</summary>
        public string? ActualVersion { get; }
        /// <summary>Gets the report containing structured conflict diagnostics.</summary>
        public TranslationReport Report { get; }
    }
}
