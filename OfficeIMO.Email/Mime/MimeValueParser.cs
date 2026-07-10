namespace OfficeIMO.Email;

internal static class MimeValueParser {
    internal static MimeValue Parse(string? input, string defaultValue, IList<EmailDiagnostic> diagnostics, string location) {
        if (string.IsNullOrWhiteSpace(input)) return new MimeValue(defaultValue);
        List<string> segments = Split(input!);
        MimeValue result = new MimeValue(segments[0].Trim().ToLowerInvariant());
        Dictionary<string, SortedDictionary<int, string>> continuations = new Dictionary<string, SortedDictionary<int, string>>(StringComparer.OrdinalIgnoreCase);

        for (int i = 1; i < segments.Count; i++) {
            int equals = segments[i].IndexOf('=');
            if (equals <= 0) continue;
            string name = segments[i].Substring(0, equals).Trim();
            string value = Unquote(segments[i].Substring(equals + 1).Trim());
            bool encoded = name.EndsWith("*", StringComparison.Ordinal);
            string baseName = name.TrimEnd('*');

            int star = baseName.LastIndexOf('*');
            if (star > 0 && int.TryParse(baseName.Substring(star + 1), NumberStyles.None, CultureInfo.InvariantCulture, out int part)) {
                string continuationName = baseName.Substring(0, star);
                SortedDictionary<int, string> values;
                if (!continuations.TryGetValue(continuationName, out values!)) {
                    values = new SortedDictionary<int, string>();
                    continuations[continuationName] = values;
                }
                values[part] = encoded ? DecodeExtended(value, diagnostics, location) : value;
            } else {
                result.Parameters[baseName] = encoded ? DecodeExtended(value, diagnostics, location) : value;
            }
        }

        foreach (KeyValuePair<string, SortedDictionary<int, string>> continuation in continuations) {
            StringBuilder builder = new StringBuilder();
            int expected = 0;
            foreach (KeyValuePair<int, string> part in continuation.Value) {
                if (part.Key != expected) {
                    diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_PARAMETER_CONTINUATION_GAP",
                        string.Concat("Parameter '", continuation.Key, "' has a missing continuation segment."),
                        EmailDiagnosticSeverity.Warning, location));
                    expected = part.Key;
                }
                builder.Append(part.Value);
                expected++;
            }
            result.Parameters[continuation.Key] = builder.ToString();
        }
        return result;
    }

    private static string DecodeExtended(string value, IList<EmailDiagnostic> diagnostics, string location) {
        string charset = "utf-8";
        string payload = value;
        int first = value.IndexOf('\'');
        int second = first >= 0 ? value.IndexOf('\'', first + 1) : -1;
        if (first > 0 && second >= 0) {
            charset = value.Substring(0, first);
            payload = value.Substring(second + 1);
        }

        using (MemoryStream output = new MemoryStream(payload.Length)) {
            for (int i = 0; i < payload.Length; i++) {
                if (payload[i] == '%' && i + 2 < payload.Length && byte.TryParse(payload.Substring(i + 1, 2),
                    NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte decoded)) {
                    output.WriteByte(decoded);
                    i += 2;
                } else {
                    output.WriteByte((byte)payload[i]);
                }
            }
            return MimeTextCodec.DecodeText(output.ToArray(), charset, diagnostics, location);
        }
    }

    private static List<string> Split(string input) {
        List<string> segments = new List<string>();
        StringBuilder current = new StringBuilder();
        bool quoted = false;
        bool escaped = false;
        foreach (char character in input) {
            if (escaped) {
                current.Append(character);
                escaped = false;
            } else if (character == '\\' && quoted) {
                current.Append(character);
                escaped = true;
            } else if (character == '"') {
                current.Append(character);
                quoted = !quoted;
            } else if (character == ';' && !quoted) {
                segments.Add(current.ToString());
                current.Clear();
            } else {
                current.Append(character);
            }
        }
        segments.Add(current.ToString());
        return segments;
    }

    private static string Unquote(string value) {
        if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"') {
            return value.Substring(1, value.Length - 2).Replace("\\\"", "\"").Replace("\\\\", "\\");
        }
        return value;
    }
}
