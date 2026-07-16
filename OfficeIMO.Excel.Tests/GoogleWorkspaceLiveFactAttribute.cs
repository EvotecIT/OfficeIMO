using Xunit;

namespace OfficeIMO.Tests {
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    internal sealed class GoogleWorkspaceLiveFactAttribute : FactAttribute {
        public GoogleWorkspaceLiveFactAttribute() {
            if (!string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_RUN_GOOGLE_WORKSPACE_LIVE"), "1", StringComparison.Ordinal)) {
                Skip = "Set OFFICEIMO_RUN_GOOGLE_WORKSPACE_LIVE=1 to run Google Workspace integration tests.";
            } else if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_ACCESS_TOKEN"))) {
                Skip = "GOOGLE_WORKSPACE_ACCESS_TOKEN is required for the disposable live-test lane.";
            } else if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_FOLDER_ID"))) {
                Skip = "GOOGLE_WORKSPACE_FOLDER_ID is required so disposable files remain isolated.";
            }
        }
    }
}
