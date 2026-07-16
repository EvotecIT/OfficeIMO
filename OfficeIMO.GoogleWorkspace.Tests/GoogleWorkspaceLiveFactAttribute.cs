using System;
using Xunit;

namespace OfficeIMO.Tests {
    internal sealed class GoogleWorkspaceLiveFactAttribute : FactAttribute {
        public GoogleWorkspaceLiveFactAttribute() {
            if (!string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_RUN_GOOGLE_WORKSPACE_LIVE"), "1", StringComparison.Ordinal)) Skip = "Set OFFICEIMO_RUN_GOOGLE_WORKSPACE_LIVE=1 to run Google Workspace integration tests.";
            else if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_ACCESS_TOKEN"))) Skip = "GOOGLE_WORKSPACE_ACCESS_TOKEN is required.";
            else if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_FOLDER_ID"))) Skip = "GOOGLE_WORKSPACE_FOLDER_ID is required for disposable files.";
        }
    }
}
