using System;

namespace OfficeIMO.Drawing;

/// <summary>Validates interactive drawing targets before they reach SVG or document exporters.</summary>
internal static class OfficeDrawingLinkPolicy {
    internal static bool TryNormalize(string? value, out string uri) {
        uri = string.Empty;
        if (string.IsNullOrWhiteSpace(value)) return false;

        for (int index = 0; index < value!.Length; index++) {
            if (char.IsControl(value[index])) return false;
        }
        string candidate = value!.Trim();
        for (int index = 0; index < candidate.Length; index++) {
            char character = candidate[index];
            if (char.IsControl(character) || char.IsWhiteSpace(character) || character == '\\') return false;
        }

        // A scheme-relative target has no stable scheme until it reaches a browser.
        if (candidate.StartsWith("//", StringComparison.Ordinal)) return false;
        if (!System.Uri.TryCreate(candidate, UriKind.RelativeOrAbsolute, out System.Uri? parsed)) return false;
        if (parsed.IsAbsoluteUri) {
            string scheme = parsed.Scheme;
            if (scheme.Equals(System.Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                || scheme.Equals(System.Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) {
                if (parsed.Host.Length == 0 || parsed.UserInfo.Length != 0) return false;
            } else if (!scheme.Equals(System.Uri.UriSchemeMailto, StringComparison.OrdinalIgnoreCase)
                       && !scheme.Equals("tel", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }
        } else {
            int colon = candidate.IndexOf(':');
            int pathDelimiter = candidate.IndexOfAny(new[] { '/', '?', '#' });
            if (colon >= 0 && (pathDelimiter < 0 || colon < pathDelimiter)) {
                // Do not let malformed scheme-shaped targets fall back to relative links.
                return false;
            }
        }

        uri = candidate;
        return true;
    }
}
