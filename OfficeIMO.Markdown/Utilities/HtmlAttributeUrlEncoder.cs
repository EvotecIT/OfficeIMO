namespace OfficeIMO.Markdown;

internal static class HtmlAttributeUrlEncoder {
    internal static string Encode(string? url) {
        return Encode(url, null);
    }

    internal static string Encode(string? url, HtmlOptions? options) {
        if (string.IsNullOrEmpty(url)) {
            return string.Empty;
        }

        var value = url!;
        return System.Net.WebUtility.HtmlEncode(NormalizeUrl(
            value,
            options?.NormalizeUrlHostsToIdn != false,
            options?.PercentEncodeTildeInUrlAttributes == true));
    }

    internal static string EncodeSrcSet(string? srcSet) {
        return EncodeSrcSet(srcSet, null);
    }

    internal static string EncodeSrcSet(string? srcSet, HtmlOptions? options) {
        if (string.IsNullOrWhiteSpace(srcSet)) {
            return string.Empty;
        }

        var encodedCandidates = new System.Collections.Generic.List<string>();
        foreach (SrcSetCandidate candidate in SrcSetParser.Parse(srcSet)) {
            string encodedUrl = Encode(candidate.Url, options);
            string encodedDescriptors = System.Net.WebUtility.HtmlEncode(candidate.Descriptor);
            encodedCandidates.Add(encodedDescriptors.Length == 0 ? encodedUrl : encodedUrl + " " + encodedDescriptors);
        }

        return string.Join(", ", encodedCandidates);
    }

    private static string NormalizeUrl(string value, bool normalizeHostsToIdn, bool percentEncodeTilde) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        if (normalizeHostsToIdn && TryNormalizeAbsoluteUrlWithIdn(value, percentEncodeTilde, out var normalized)) {
            return normalized;
        }

        return PercentEncodeRelativeUrl(value, percentEncodeTilde);
    }

    private static bool TryNormalizeAbsoluteUrlWithIdn(string value, bool percentEncodeTilde, out string normalized) {
        normalized = string.Empty;
        if (!ContainsNonAscii(value)) {
            return false;
        }

        var schemeSeparator = value.IndexOf("://", System.StringComparison.Ordinal);
        if (schemeSeparator <= 0) {
            return false;
        }

        var scheme = value.Substring(0, schemeSeparator);
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
        builder.Append(PercentEncodeRelativeUrl(value.Substring(authorityEnd), percentEncodeTilde));
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

    private static string PercentEncodeRelativeUrl(string value, bool percentEncodeTilde) {
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
                if (ShouldPercentEncodeAsciiUrlCharacter(current, percentEncodeTilde)) {
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

    private static bool ShouldPercentEncodeAsciiUrlCharacter(char value, bool percentEncodeTilde) =>
        value <= 0x20
        || (percentEncodeTilde && value == '~')
        || value == '\''
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
