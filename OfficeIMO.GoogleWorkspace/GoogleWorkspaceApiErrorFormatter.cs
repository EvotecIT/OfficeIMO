using System.Text.Json;

namespace OfficeIMO.GoogleWorkspace {
    public static class GoogleWorkspaceApiErrorFormatter {
        public static string? Format(string responseBody) {
            if (string.IsNullOrWhiteSpace(responseBody)) {
                return null;
            }

            JsonElement root;
            try {
                using JsonDocument document = JsonDocument.Parse(responseBody);
                root = document.RootElement.Clone();
            } catch (JsonException) {
                return null;
            }

            if (!root.TryGetProperty("error", out JsonElement error)) {
                return null;
            }

            if (error.ValueKind == JsonValueKind.String) {
                string? code = error.GetString();
                if (string.IsNullOrWhiteSpace(code)) return null;
                string? description = root.TryGetProperty("error_description", out JsonElement descriptionElement)
                    && descriptionElement.ValueKind == JsonValueKind.String
                    ? descriptionElement.GetString()
                    : null;
                return string.IsNullOrWhiteSpace(description)
                    ? code
                    : code + " (" + description + ")";
            }

            if (error.ValueKind != JsonValueKind.Object) return null;

            var parts = new List<string>();
            string? status = ReadString(error, "status");
            if (!string.IsNullOrWhiteSpace(status)) {
                parts.Add(status!);
            }

            string? message = ReadString(error, "message");
            if (!string.IsNullOrWhiteSpace(message)) {
                parts.Add(message!);
            }

            string? reason = ReadFirstErrorValue(error, "reason");
            if (!string.IsNullOrWhiteSpace(reason)) {
                parts.Add("reason=" + reason);
            }

            string? domain = ReadFirstErrorValue(error, "domain");
            if (!string.IsNullOrWhiteSpace(domain)) {
                parts.Add("domain=" + domain);
            }

            if (parts.Count == 0) {
                return null;
            }

            return string.Join("; ", parts);
        }

        private static string? ReadString(JsonElement parent, string propertyName) {
            return parent.TryGetProperty(propertyName, out JsonElement value)
                && value.ValueKind == JsonValueKind.String
                ? value.GetString()
                : null;
        }

        private static string? ReadFirstErrorValue(JsonElement error, string propertyName) {
            if (!error.TryGetProperty("errors", out JsonElement errors) || errors.ValueKind != JsonValueKind.Array) {
                return null;
            }

            foreach (JsonElement item in errors.EnumerateArray()) {
                if (item.ValueKind != JsonValueKind.Object) continue;
                string? value = ReadString(item, propertyName);
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }

            return null;
        }
    }
}
