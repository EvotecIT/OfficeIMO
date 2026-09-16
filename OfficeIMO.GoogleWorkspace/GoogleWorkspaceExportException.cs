namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Describes the high-level category of a Google Workspace export failure.
    /// </summary>
    public enum GoogleWorkspaceFailureKind {
        /// <summary>OAuth token acquisition or validation failed.</summary>
        TokenAcquisition = 0,
        /// <summary>Token acquisition failed with evidence that points to service-account domain-wide delegation.</summary>
        DomainWideDelegation = 1,
        /// <summary>A Google API request failed.</summary>
        ApiRequest = 2,
        /// <summary>A Google API request exceeded its configured timeout.</summary>
        RequestTimeout = 3,
        /// <summary>The caller canceled the operation.</summary>
        Canceled = 4,
        /// <summary>A mutation may have committed remotely but no conclusive response was received.</summary>
        AmbiguousMutation = 5,
    }

    /// <summary>
    /// Export failure that preserves the translation report collected before the operation failed.
    /// </summary>
    public sealed class GoogleWorkspaceExportException : Exception {
        /// <summary>Creates an export failure with its structured category and accumulated translation report.</summary>
        /// <param name="message">Human-readable failure description.</param>
        /// <param name="failureKind">High-level failure category.</param>
        /// <param name="report">Translation report available when the operation failed.</param>
        /// <param name="innerException">Underlying authentication, transport, or API exception.</param>
        public GoogleWorkspaceExportException(
            string message,
            GoogleWorkspaceFailureKind failureKind,
            TranslationReport report,
            Exception innerException)
            : base(message, innerException) {
            FailureKind = failureKind;
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        /// <summary>Gets the high-level failure category.</summary>
        public GoogleWorkspaceFailureKind FailureKind { get; }
        /// <summary>Gets the translation report available when the operation failed.</summary>
        public TranslationReport Report { get; }
    }

    /// <summary>
    /// Export cancellation that preserves the translation report collected before cancellation.
    /// </summary>
    public sealed class GoogleWorkspaceExportCanceledException : OperationCanceledException {
        /// <summary>Creates a cancellation failure while preserving the caller's cancellation token.</summary>
        /// <param name="message">Human-readable cancellation description.</param>
        /// <param name="report">Translation report available when cancellation was observed.</param>
        /// <param name="innerException">Original cancellation exception.</param>
        public GoogleWorkspaceExportCanceledException(
            string message,
            TranslationReport report,
            OperationCanceledException innerException)
            : base(message, innerException, innerException.CancellationToken) {
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        /// <summary>Gets the fixed <see cref="GoogleWorkspaceFailureKind.Canceled"/> category.</summary>
        public GoogleWorkspaceFailureKind FailureKind => GoogleWorkspaceFailureKind.Canceled;
        /// <summary>Gets the translation report available when cancellation was observed.</summary>
        public TranslationReport Report { get; }
    }

    /// <summary>Creates consistent structured exceptions and diagnostics for Workspace export failures.</summary>
    public static class GoogleWorkspaceFailureDiagnostics {
        /// <summary>Classifies a credential failure, records it in the report, and returns an export exception.</summary>
        /// <param name="operationName">User-facing name of the export operation.</param>
        /// <param name="scopes">OAuth scopes requested from the credential source.</param>
        /// <param name="session">Session whose credential policy was applied.</param>
        /// <param name="report">Report to receive the failure diagnostic.</param>
        /// <param name="exception">Underlying credential exception.</param>
        /// <returns>A domain-delegation or token-acquisition export exception.</returns>
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
                    failureKind: GoogleWorkspaceFailureKind.DomainWideDelegation,
                    code: "WORKSPACE.AUTH.DOMAIN_DELEGATION_FAILED",
                    action: TranslationAction.Fail);

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
                failureKind: GoogleWorkspaceFailureKind.TokenAcquisition,
                code: GoogleWorkspaceDiagnosticCodes.AuthenticationFailed,
                action: TranslationAction.Fail);

            return new GoogleWorkspaceExportException(
                messageWithScopes,
                GoogleWorkspaceFailureKind.TokenAcquisition,
                report,
                exception);
        }

        /// <summary>Records a general Google API failure and returns its structured export exception.</summary>
        /// <param name="operationName">User-facing name of the export operation.</param>
        /// <param name="sessionOptions">Session diagnostic settings, when available.</param>
        /// <param name="report">Report to receive the failure diagnostic.</param>
        /// <param name="exception">Underlying API or transport exception.</param>
        /// <returns>An API-request export exception.</returns>
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
                failureKind: GoogleWorkspaceFailureKind.ApiRequest,
                code: GoogleWorkspaceDiagnosticCodes.RequestFailed,
                action: TranslationAction.Fail);

            return new GoogleWorkspaceExportException(
                message,
                GoogleWorkspaceFailureKind.ApiRequest,
                report,
                exception);
        }

        /// <summary>Records a request timeout and returns its structured export exception.</summary>
        /// <param name="operationName">User-facing name of the export operation.</param>
        /// <param name="sessionOptions">Session diagnostic settings, when available.</param>
        /// <param name="report">Report to receive the timeout diagnostic.</param>
        /// <param name="exception">Timeout exception raised by the request.</param>
        /// <returns>A request-timeout export exception.</returns>
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
                failureKind: GoogleWorkspaceFailureKind.RequestTimeout,
                code: GoogleWorkspaceDiagnosticCodes.RequestTimedOut,
                action: TranslationAction.Fail);

            return new GoogleWorkspaceExportException(
                message,
                GoogleWorkspaceFailureKind.RequestTimeout,
                report,
                exception);
        }

        /// <summary>
        /// Creates an actionable export failure for a mutation that must be reconciled before retry.
        /// </summary>
        public static GoogleWorkspaceExportException CreateAmbiguousMutationFailure(
            string operationName,
            GoogleWorkspaceSessionOptions? sessionOptions,
            TranslationReport report,
            GoogleWorkspaceAmbiguousMutationException exception) {
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            string requestId = string.IsNullOrWhiteSpace(exception.Receipt.RequestId)
                ? "not supplied"
                : exception.Receipt.RequestId!;
            string message = $"{operationName} may have committed remotely without returning a response. Reconcile target '{exception.Receipt.Target}' and request id '{requestId}' before retrying.";

            GoogleWorkspaceDiagnosticsDispatcher.AddUnique(
                report,
                sessionOptions,
                TranslationSeverity.Error,
                "AmbiguousMutation",
                message,
                failureKind: GoogleWorkspaceFailureKind.AmbiguousMutation,
                code: GoogleWorkspaceDiagnosticCodes.AmbiguousMutation,
                action: TranslationAction.Fail);

            return new GoogleWorkspaceExportException(
                message,
                GoogleWorkspaceFailureKind.AmbiguousMutation,
                report,
                exception);
        }

        /// <summary>Records caller cancellation and returns an exception that preserves the original cancellation token.</summary>
        /// <param name="operationName">User-facing name of the export operation.</param>
        /// <param name="sessionOptions">Session diagnostic settings, when available.</param>
        /// <param name="report">Report to receive the cancellation diagnostic.</param>
        /// <param name="exception">Original cancellation exception.</param>
        /// <returns>A structured cancellation exception.</returns>
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
                failureKind: GoogleWorkspaceFailureKind.Canceled,
                code: GoogleWorkspaceDiagnosticCodes.RequestCanceled,
                action: TranslationAction.Fail);

            return new GoogleWorkspaceExportCanceledException(
                message,
                report,
                exception);
        }

        private static bool IsDomainWideDelegationFailure(GoogleWorkspaceSession session, Exception exception) {
            if (!session.Options.UseDomainWideDelegation
                || string.IsNullOrWhiteSpace(session.Options.SubjectUser)) {
                return false;
            }

            string diagnostic = BuildDiagnostic(exception);
            return diagnostic.Contains("unauthorized_client", StringComparison.OrdinalIgnoreCase)
                || diagnostic.Contains("domain-wide delegation", StringComparison.OrdinalIgnoreCase)
                || diagnostic.Contains("domain wide delegation", StringComparison.OrdinalIgnoreCase)
                || diagnostic.Contains("delegation denied", StringComparison.OrdinalIgnoreCase)
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
