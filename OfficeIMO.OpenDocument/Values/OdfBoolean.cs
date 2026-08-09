namespace OfficeIMO.OpenDocument;

internal static class OdfBoolean {
    internal static bool TryParseXml(string? lexical, out bool value) {
        if (lexical == "true" || lexical == "1") {
            value = true;
            return true;
        }
        if (lexical == "false" || lexical == "0") {
            value = false;
            return true;
        }
        value = false;
        return false;
    }

    internal static bool TryParseCompatible(string? lexical, out bool value) {
        if (TryParseXml(lexical, out value)) return true;
        if (string.Equals(lexical, "true", StringComparison.OrdinalIgnoreCase)) {
            value = true;
            return true;
        }
        if (string.Equals(lexical, "false", StringComparison.OrdinalIgnoreCase)) {
            value = false;
            return true;
        }
        value = false;
        return false;
    }

    internal static bool ReadCompatible(string? lexical, bool fallback) =>
        TryParseCompatible(lexical, out bool value) ? value : fallback;
}
