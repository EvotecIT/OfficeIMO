namespace OfficeIMO.OpenDocument;

/// <summary>Provides canonical parsing for URI references used by OpenDocument content.</summary>
public static class OdfUriReference {
    /// <summary>Decodes a non-empty same-document fragment reference.</summary>
    /// <param name="href">The URI reference, including its leading <c>#</c>.</param>
    /// <param name="fragment">The decoded fragment identifier when parsing succeeds.</param>
    /// <returns><see langword="true"/> when <paramref name="href"/> is a valid non-empty fragment reference.</returns>
    public static bool TryDecodeFragment(string? href, out string fragment) {
        fragment = string.Empty;
        if (href == null || href.Length == 0 || href[0] != '#' || href.Length == 1) return false;
        string encoded = href.Substring(1);
        for (int index = 0; index < encoded.Length; index++) {
            if (encoded[index] != '%') continue;
            if (index + 2 >= encoded.Length || !IsHex(encoded[index + 1]) || !IsHex(encoded[index + 2])) return false;
            index += 2;
        }
        try {
            fragment = Uri.UnescapeDataString(encoded);
            return fragment.Length > 0;
        } catch (UriFormatException) {
            fragment = string.Empty;
            return false;
        }
    }

    private static bool IsHex(char value) =>
        value >= '0' && value <= '9' || value >= 'A' && value <= 'F' || value >= 'a' && value <= 'f';
}
