namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Raised before mutation when a remote resource no longer matches the caller's observed version.
    /// </summary>
    public sealed class GoogleWorkspaceConflictException : InvalidOperationException {
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

        public string TargetId { get; }
        public string? ExpectedVersion { get; }
        public string? ActualVersion { get; }
        public TranslationReport Report { get; }
    }
}
