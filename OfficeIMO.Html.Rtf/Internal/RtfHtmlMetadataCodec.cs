namespace OfficeIMO.Html;

internal static class RtfHtmlMetadataCodec {
    internal static string Encode(IReadOnlyDictionary<string, string> values) {
        var builder = new StringBuilder();
        foreach (KeyValuePair<string, string> pair in values.OrderBy(pair => pair.Key, StringComparer.Ordinal)) {
            if (string.IsNullOrEmpty(pair.Key)) {
                continue;
            }

            builder.Append(pair.Key);
            builder.Append('=');
            builder.Append(Convert.ToBase64String(Encoding.UTF8.GetBytes(pair.Value ?? string.Empty)));
            builder.Append('\n');
        }

        return Convert.ToBase64String(Encoding.UTF8.GetBytes(builder.ToString()));
    }

    internal static Dictionary<string, string> Decode(string? value) {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(value)) {
            return values;
        }

        string text;
        try {
            text = Encoding.UTF8.GetString(Convert.FromBase64String(value!));
        } catch (FormatException) {
            return values;
        }

        string[] lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string line in lines) {
            int separator = line.IndexOf('=');
            if (separator <= 0) {
                continue;
            }

            string key = line.Substring(0, separator).Trim();
            string encoded = line.Substring(separator + 1).Trim();
            if (key.Length == 0) {
                continue;
            }

            try {
                values[key] = Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
            } catch (FormatException) {
                // Ignore malformed entries while preserving the rest of the metadata payload.
            }
        }

        return values;
    }
}
