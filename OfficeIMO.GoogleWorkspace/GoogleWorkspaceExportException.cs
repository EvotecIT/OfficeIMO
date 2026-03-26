namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Describes the high-level category of a Google Workspace export failure.
    /// </summary>
    public enum GoogleWorkspaceFailureKind {
        TokenAcquisition = 0,
        DomainWideDelegation = 1,
        ApiRequest = 2,
        RequestTimeout = 3,
        Canceled = 4,
    }

    /// <summary>
    /// Export failure that preserves the translation report collected before the operation failed.
    /// </summary>
    public sealed class GoogleWorkspaceExportException : Exception {
        public GoogleWorkspaceExportException(
            string message,
            GoogleWorkspaceFailureKind failureKind,
            TranslationReport report,
            Exception innerException)
            : base(message, innerException) {
            FailureKind = failureKind;
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        public GoogleWorkspaceFailureKind FailureKind { get; }
        public TranslationReport Report { get; }
    }

    /// <summary>
    /// Export cancellation that preserves the translation report collected before cancellation.
    /// </summary>
    public sealed class GoogleWorkspaceExportCanceledException : OperationCanceledException {
        public GoogleWorkspaceExportCanceledException(
            string message,
            TranslationReport report,
            OperationCanceledException innerException)
            : base(message, innerException, innerException.CancellationToken) {
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        public GoogleWorkspaceFailureKind FailureKind => GoogleWorkspaceFailureKind.Canceled;
        public TranslationReport Report { get; }
    }

    public static class GoogleWorkspaceFailureDiagnostics {
        public static GoogleWorkspaceExportException CreateTokenAcquisitionFailure(
            string operationName,
            IReadOnlyList<string> scopes,
            GoogleWorkspaceSession session,
            TranslationReport report,
            Exception exception) {
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (session == null) throw new ArgumentNullException(nameof(session));
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            if (IsDomainWideDelegationFailure(session, exception)) {
                string delegatedUser = string.IsNullOrWhiteSpace(session.Options.SubjectUser)
                    ? "the configured delegated user"
                    : $"delegated user '{session.Options.SubjectUser}'";
                string message = $"{operationName} could not acquire a delegated Google access token for {delegatedUser}. The service account may be missing domain-wide delegation approval or the delegated user may be invalid. Original error: {exception.Message}";

                GoogleWorkspaceDiagnosticsDispatcher.AddUnique(
                    report,
                    session.Options,
                    TranslationSeverity.Error,
                    "DomainWideDelegation",
                    message,
                    failureKind: GoogleWorkspaceFailureKind.DomainWideDelegation);

                return new GoogleWorkspaceExportException(
                    message,
                    GoogleWorkspaceFailureKind.DomainWideDelegation,
                    report,
                    exception);
            }

            string requestedScopes = scopes == null || scopes.Count == 0
                ? "<none>"
                : string.Join(", ", scopes);
            string messageWithScopes = $"{operationName} could not acquire a Google access token from {session.CredentialSource.GetType().Name} for scopes [{requestedScopes}]. Original error: {exception.Message}";

            GoogleWorkspaceDiagnosticsDispatcher.AddUnique(
                report,
                session.Options,
                TranslationSeverity.Error,
                "Authentication",
                messageWithScopes,
                failureKind: GoogleWorkspaceFailureKind.TokenAcquisition);

            return new GoogleWorkspaceExportException(
                messageWithScopes,
                GoogleWorkspaceFailureKind.TokenAcquisition,
                report,
                exception);
        }

        public static GoogleWorkspaceExportException CreateApiFailure(
            string operationName,
            GoogleWorkspaceSessionOptions? sessionOptions,
            TranslationReport report,
            Exception exception) {
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            string message = $"{operationName} failed during Google API execution. Original error: {exception.Message}";

            GoogleWorkspaceDiagnosticsDispatcher.AddUnique(
                report,
                sessionOptions,
                TranslationSeverity.Error,
                "ApiFailures",
                message,
                failureKind: GoogleWorkspaceFailureKind.ApiRequest);

            return new GoogleWorkspaceExportException(
                message,
                GoogleWorkspaceFailureKind.ApiRequest,
                report,
                exception);
        }

        public static GoogleWorkspaceExportException CreateRequestTimeoutFailure(
            string operationName,
            GoogleWorkspaceSessionOptions? sessionOptions,
            TranslationReport report,
            TaskCanceledException exception) {
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            string message = $"{operationName} timed out while waiting for Google service communication to complete. Original error: {exception.Message}";

            GoogleWorkspaceDiagnosticsDispatcher.AddUnique(
                report,
                sessionOptions,
                TranslationSeverity.Error,
                "RequestTimeout",
                message,
                failureKind: GoogleWorkspaceFailureKind.RequestTimeout);

            return new GoogleWorkspaceExportException(
                message,
                GoogleWorkspaceFailureKind.RequestTimeout,
                report,
                exception);
        }

        public static GoogleWorkspaceExportCanceledException CreateCanceledFailure(
            string operationName,
            GoogleWorkspaceSessionOptions? sessionOptions,
            TranslationReport report,
            OperationCanceledException exception) {
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            string message = $"{operationName} was canceled by the caller before the Google export completed.";

            GoogleWorkspaceDiagnosticsDispatcher.AddUnique(
                report,
                sessionOptions,
                TranslationSeverity.Warning,
                "Cancellation",
                message,
                failureKind: GoogleWorkspaceFailureKind.Canceled);

            return new GoogleWorkspaceExportCanceledException(
                message,
                report,
                exception);
        }

        private static bool IsDomainWideDelegationFailure(GoogleWorkspaceSession session, Exception exception) {
            if (!session.Options.UseDomainWideDelegation) {
                return false;
            }

            string diagnostic = BuildDiagnostic(exception);
            return diagnostic.Contains("unauthorized_client", StringComparison.OrdinalIgnoreCase)
                || diagnostic.Contains("domain-wide delegation", StringComparison.OrdinalIgnoreCase)
                || diagnostic.Contains("domain wide delegation", StringComparison.OrdinalIgnoreCase)
                || diagnostic.Contains("delegation denied", StringComparison.OrdinalIgnoreCase)
                || diagnostic.Contains("delegation", StringComparison.OrdinalIgnoreCase)
                || diagnostic.Contains("invalid_grant", StringComparison.OrdinalIgnoreCase)
                || diagnostic.Contains("not a valid email", StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildDiagnostic(Exception exception) {
            if (exception.InnerException == null) {
                return exception.Message ?? string.Empty;
            }

            return (exception.Message ?? string.Empty) + Environment.NewLine + BuildDiagnostic(exception.InnerException);
        }
    }
}
