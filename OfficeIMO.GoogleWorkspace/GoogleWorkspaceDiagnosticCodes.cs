using System.Text;

namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Stable diagnostic codes emitted by the shared Google Workspace layer.
    /// </summary>
    public static class GoogleWorkspaceDiagnosticCodes {
        public const string ApiRetry = "WORKSPACE.API.RETRY";
        public const string AuthenticationFailed = "WORKSPACE.AUTH.FAILED";
        public const string RequestCanceled = "WORKSPACE.REQUEST.CANCELED";
        public const string RequestFailed = "WORKSPACE.REQUEST.FAILED";
        public const string RequestTimedOut = "WORKSPACE.REQUEST.TIMED_OUT";

        public static string Resolve(string? code, string? feature) {
            if (!string.IsNullOrWhiteSpace(code)) {
                return code!.Trim().ToUpperInvariant();
            }

            if (string.IsNullOrWhiteSpace(feature)) {
                return "WORKSPACE.GENERAL";
            }

            var builder = new StringBuilder("WORKSPACE.");
            bool previousWasSeparator = true;
            foreach (char character in feature!) {
                if (char.IsLetterOrDigit(character)) {
                    if (char.IsUpper(character) && !previousWasSeparator && builder.Length > 0 && builder[builder.Length - 1] != '.') {
                        builder.Append('_');
                    }

                    builder.Append(char.ToUpperInvariant(character));
                    previousWasSeparator = false;
                } else if (!previousWasSeparator) {
                    builder.Append('_');
                    previousWasSeparator = true;
                }
            }

            while (builder.Length > 0 && builder[builder.Length - 1] == '_') {
                builder.Length--;
            }

            return builder.ToString();
        }
    }
}
