using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Examples.Google {
    internal static class GoogleWorkspaceExampleSession {
        public static GoogleWorkspaceSession? TryCreateSession() {
            string? serviceAccountJsonPath = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_SERVICE_ACCOUNT_JSON_PATH");
            string? serviceAccountJson = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_SERVICE_ACCOUNT_JSON");
            string? accessToken = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_ACCESS_TOKEN");

            var sessionOptions = new GoogleWorkspaceSessionOptions {
                ApplicationName = "OfficeIMO.Examples",
                DefaultDriveId = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_DRIVE_ID"),
                DefaultFolderId = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_FOLDER_ID"),
                SubjectUser = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_SUBJECT_USER"),
                UseDomainWideDelegation = !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_SUBJECT_USER")),
                MaxRetryCount = 5,
                RetryBaseDelay = TimeSpan.FromMilliseconds(250),
                RetryMaxDelay = TimeSpan.FromSeconds(10),
                RequestTimeout = TimeSpan.FromSeconds(120),
            };

            if (!string.IsNullOrWhiteSpace(serviceAccountJsonPath)) {
                return new GoogleWorkspaceSession(
                    GoogleServiceAccountCredentialSource.FromFile(serviceAccountJsonPath, sessionOptions),
                    sessionOptions);
            }

            if (!string.IsNullOrWhiteSpace(serviceAccountJson)) {
                return new GoogleWorkspaceSession(
                    GoogleServiceAccountCredentialSource.FromJson(serviceAccountJson, sessionOptions),
                    sessionOptions);
            }

            if (string.IsNullOrWhiteSpace(accessToken)) {
                return null;
            }

            return new GoogleWorkspaceSession(
                new StaticAccessTokenCredentialSource(accessToken),
                sessionOptions);
        }

        public static void PrintMissingTokenMessage() {
            Console.WriteLine("Skipping Google export because no Google credential source environment variable is set.");
            Console.WriteLine("Use GOOGLE_WORKSPACE_ACCESS_TOKEN, GOOGLE_WORKSPACE_SERVICE_ACCOUNT_JSON, or GOOGLE_WORKSPACE_SERVICE_ACCOUNT_JSON_PATH.");
            Console.WriteLine("Optional environment variables: GOOGLE_WORKSPACE_FOLDER_ID, GOOGLE_WORKSPACE_DRIVE_ID, and GOOGLE_WORKSPACE_SUBJECT_USER.");
        }

        public static void PrintExportFailure(GoogleWorkspaceExportException exception) {
            if (exception == null) throw new ArgumentNullException(nameof(exception));

            foreach (var entry in exception.ToDiagnosticEntries()) {
                string failureKind = entry.FailureKind.HasValue
                    ? $" [{entry.FailureKind.Value}]"
                    : string.Empty;
                string path = string.IsNullOrWhiteSpace(entry.Path)
                    ? string.Empty
                    : $" ({entry.Path})";

                Console.WriteLine($"  {entry.Severity}: {entry.Feature}{failureKind} - {entry.Message}{path}");
            }
        }
    }
}
