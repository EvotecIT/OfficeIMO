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

        if (TryNormalizeAbsoluteUrlWithIdn(value, out var normalized)) {
            return normalized;
        }

        return PercentEncodeRelativeUrl(value);
    }

    private static bool TryNormalizeAbsoluteUrlWithIdn(string value, out string normalized) {
        normalized = string.Empty;
        if (!ContainsNonAscii(value)) {
            return false;
        }

        var schemeSeparator = value.IndexOf("://", System.StringComparison.Ordinal);
        if (schemeSeparator <= 0) {
            return false;
        }

        var scheme = value.Substring(0, schemeSeparator);
        if (!string.Equals(scheme, "http", System.StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(scheme, "https", System.StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        var builder = new System.Text.StringBuilder(value.Length);
        var authorityStart = schemeSeparator + 3;
        var authorityEnd = value.Length;
        for (var i = authorityStart; i < value.Length; i++) {
            var current = value[i];
            if (current == '/' || current == '?' || current == '#') {
                authorityEnd = i;
                break;
            }
        }

        var authority = value.Substring(authorityStart, authorityEnd - authorityStart);
        if (authority.Length == 0) {
            return false;
        }

        builder.Append(scheme).Append("://").Append(NormalizeAuthorityHost(authority));
        builder.Append(PercentEncodeRelativeUrl(value.Substring(authorityEnd)));
        normalized = builder.ToString();
        return true;
    }

    private static string NormalizeAuthorityHost(string authority) {
        var userInfoEnd = authority.LastIndexOf('@');
        var hostStart = userInfoEnd >= 0 ? userInfoEnd + 1 : 0;
        var hostEnd = authority.Length;

        if (hostStart < authority.Length && authority[hostStart] == '[') {
            var bracketEnd = authority.IndexOf(']', hostStart + 1);
            if (bracketEnd >= 0) {
                hostEnd = bracketEnd + 1;
            }
        } else {
            var colon = authority.LastIndexOf(':');
            if (colon > hostStart) {
                hostEnd = colon;
            }
        }

        var prefix = authority.Substring(0, hostStart);
        var host = authority.Substring(hostStart, hostEnd - hostStart);
        var suffix = authority.Substring(hostEnd);
        if (host.Length == 0 || host[0] == '[' || !ContainsNonAscii(host)) {
            return authority;
        }

        var labels = host.Split('.');
        for (var i = 0; i < labels.Length; i++) {
            if (!ContainsNonAscii(labels[i])) {
                continue;
            }

            try {
                labels[i] = new System.Globalization.IdnMapping().GetAscii(labels[i]);
            } catch {
                return authority;
            }
        }

        return prefix + string.Join(".", labels) + suffix;
    }

    private static bool ContainsNonAscii(string value) {
        for (var i = 0; i < value.Length; i++) {
            if (value[i] > 0x7F) {
                return true;
            }
        }

        return false;
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
                if (ShouldPercentEncodeAsciiUrlCharacter(current)) {
                    AppendPercentEncodedByte(builder, (byte)current);
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
                AppendPercentEncodedByte(builder, bytes[b]);
            }
        }

        return builder.ToString();
    }

    private static bool ShouldPercentEncodeAsciiUrlCharacter(char value) =>
        value <= 0x20
        || value == '"'
        || value == '<'
        || value == '>'
        || value == '\\'
        || value == '['
        || value == ']'
        || value == '^'
        || value == '`'
        || value == '{'
        || value == '|'
        || value == '}';

    private static void AppendPercentEncodedByte(System.Text.StringBuilder builder, byte value) {
        builder.Append('%');
        builder.Append(value.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
    }

    private static bool IsHex(char value) =>
        (value >= '0' && value <= '9')
        || (value >= 'A' && value <= 'F')
        || (value >= 'a' && value <= 'f');
}
