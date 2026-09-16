using System.Text;

namespace OfficeIMO.GoogleWorkspace {
    /// <summary>
    /// Stable diagnostic codes emitted by the shared Google Workspace layer.
    /// </summary>
    public static class GoogleWorkspaceDiagnosticCodes {
        /// <summary>Identifies a retry of a transient Google Workspace API request.</summary>
        public const string ApiRetry = "WORKSPACE.API.RETRY";
        /// <summary>Identifies a mutation whose remote outcome cannot be determined safely.</summary>
        public const string AmbiguousMutation = "WORKSPACE.MUTATION.OUTCOME_AMBIGUOUS";
        /// <summary>Identifies failure to acquire or validate Google credentials.</summary>
        public const string AuthenticationFailed = "WORKSPACE.AUTH.FAILED";
        /// <summary>Identifies a request canceled by the caller.</summary>
        public const string RequestCanceled = "WORKSPACE.REQUEST.CANCELED";
        /// <summary>Identifies a Google Workspace request that failed without a more specific diagnostic.</summary>
        public const string RequestFailed = "WORKSPACE.REQUEST.FAILED";
        /// <summary>Identifies a Google Workspace request that exceeded its configured timeout.</summary>
        public const string RequestTimedOut = "WORKSPACE.REQUEST.TIMED_OUT";

        /// <summary>Returns an explicit diagnostic code or derives a stable Workspace-prefixed code from a feature name.</summary>
        /// <param name="code">Explicit code to normalize; takes precedence when nonempty.</param>
        /// <param name="feature">Feature name used when <paramref name="code"/> is absent.</param>
        /// <returns>An uppercase diagnostic code, or <c>WORKSPACE.GENERAL</c> when neither input is supplied.</returns>
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
