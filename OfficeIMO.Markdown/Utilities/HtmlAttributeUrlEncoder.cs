namespace OfficeIMO.Markdown;

internal static class HtmlAttributeUrlEncoder {
    internal static string Encode(string? url) {
        if (string.IsNullOrEmpty(url)) {
            return string.Empty;
        }

        var value = url!;
        return System.Net.WebUtility.HtmlEncode(NormalizeUrl(value));
    }

    internal static string EncodeSrcSet(string? srcSet) {
        if (string.IsNullOrWhiteSpace(srcSet)) {
            return string.Empty;
        }

        var encodedCandidates = new System.Collections.Generic.List<string>();
        foreach (SrcSetCandidate candidate in SrcSetParser.Parse(srcSet)) {
            string encodedUrl = Encode(candidate.Url);
            string encodedDescriptors = System.Net.WebUtility.HtmlEncode(candidate.Descriptor);
            encodedCandidates.Add(encodedDescriptors.Length == 0 ? encodedUrl : encodedUrl + " " + encodedDescriptors);
        }

        return string.Join(", ", encodedCandidates);
    }

    private static string NormalizeUrl(string value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        return PercentEncodeRelativeUrl(value);
    }

    private static string PercentEncodeRelativeUrl(string value) {
        var builder = new System.Text.StringBuilder(value.Length);
        for (var i = 0; i < value.Length; i++) {
            var current = value[i];
            if (current == '%' && i + 2 < value.Length && IsHex(value[i + 1]) && IsHex(value[i + 2])) {
                builder.Append('%');
                builder.Append(value[i + 1]);
                builder.Append(value[i + 2]);
                i += 2;
                continue;
            }

            if (current <= 0x7F) {
                if (current == ' ') {
                    builder.Append("%20");
                } else {
                    builder.Append(current);
                }
                continue;
            }

            string scalar;
            if (char.IsHighSurrogate(current) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1])) {
                scalar = value.Substring(i, 2);
                i++;
            } else {
                scalar = current.ToString();
            }

            var bytes = System.Text.Encoding.UTF8.GetBytes(scalar);
            for (var b = 0; b < bytes.Length; b++) {
                builder.Append('%');
                builder.Append(bytes[b].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        return builder.ToString();
    }

    private static bool IsHex(char value) =>
        (value >= '0' && value <= '9')
        || (value >= 'A' && value <= 'F')
        || (value >= 'a' && value <= 'f');
}
