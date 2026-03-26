using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.GoogleWorkspace {
    public static class GoogleWorkspaceApiErrorFormatter {
        public static string? Format(string responseBody) {
            if (string.IsNullOrWhiteSpace(responseBody)) {
                return null;
            }

            var oauthError = TryDeserialize<GoogleOAuthErrorEnvelope>(responseBody);
            if (oauthError != null && !string.IsNullOrWhiteSpace(oauthError.Error)) {
                return string.IsNullOrWhiteSpace(oauthError.ErrorDescription)
                    ? oauthError.Error
                    : oauthError.Error + " (" + oauthError.ErrorDescription + ")";
            }

            var apiError = TryDeserialize<GoogleApiErrorEnvelope>(responseBody);
            if (apiError?.Error == null) {
                return null;
            }

            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(apiError.Error.Status)) {
                parts.Add(apiError.Error.Status!);
            }

            if (!string.IsNullOrWhiteSpace(apiError.Error.Message)) {
                parts.Add(apiError.Error.Message!);
            }

            var reason = apiError.Error.Errors?
                .Select(error => error.Reason)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            if (!string.IsNullOrWhiteSpace(reason)) {
                parts.Add("reason=" + reason);
            }

            var domain = apiError.Error.Errors?
                .Select(error => error.Domain)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            if (!string.IsNullOrWhiteSpace(domain)) {
                parts.Add("domain=" + domain);
            }

            if (parts.Count == 0) {
                return null;
            }

            return string.Join("; ", parts);
        }

        private static T? TryDeserialize<T>(string json) where T : class {
            try {
                return JsonSerializer.Deserialize<T>(json);
            } catch (JsonException) {
                return null;
            }
        }

        private sealed class GoogleOAuthErrorEnvelope {
            [JsonPropertyName("error")]
            public string? Error { get; set; }

            [JsonPropertyName("error_description")]
            public string? ErrorDescription { get; set; }
        }

        private sealed class GoogleApiErrorEnvelope {
            [JsonPropertyName("error")]
            public GoogleApiErrorBody? Error { get; set; }
        }

        private sealed class GoogleApiErrorBody {
            [JsonPropertyName("message")]
            public string? Message { get; set; }

            [JsonPropertyName("status")]
            public string? Status { get; set; }

            [JsonPropertyName("errors")]
            public List<GoogleApiErrorItem>? Errors { get; set; }
        }

        private sealed class GoogleApiErrorItem {
            [JsonPropertyName("reason")]
            public string? Reason { get; set; }

            [JsonPropertyName("domain")]
            public string? Domain { get; set; }
        }
    }
}
