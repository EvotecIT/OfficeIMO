namespace OfficeIMO.Html;

/// <summary>
/// Shared conversion support rules for picture source content types.
/// </summary>
internal static class HtmlPictureSourceSupport {
    /// <summary>
    /// Returns whether a picture source type can be consumed by the shared OfficeIMO conversion adapters.
    /// </summary>
    internal static bool IsSupportedConversionContentType(string? type) {
        if (type == null) {
            return true;
        }

        string normalized = TrimAsciiWhitespace(type);
        if (normalized.Length == 0) return true;
        int parameterStart = normalized.IndexOf(';');
        if (parameterStart >= 0) {
            normalized = TrimAsciiWhitespace(normalized.Substring(0, parameterStart));
        }

        switch (normalized.ToLowerInvariant()) {
            case "image/bmp":
            case "image/gif":
            case "image/jpeg":
            case "image/jpg":
            case "image/png":
            case "image/svg+xml":
            case "image/webp":
            case "image/x-icon":
            case "image/vnd.microsoft.icon":
                return true;
            default:
                return false;
        }
    }

    private static string TrimAsciiWhitespace(string value) {
        int start = 0;
        while (start < value.Length && IsAsciiWhitespace(value[start])) start++;
        int end = value.Length;
        while (end > start && IsAsciiWhitespace(value[end - 1])) end--;
        return start == 0 && end == value.Length ? value : value.Substring(start, end - start);
    }

    private static bool IsAsciiWhitespace(char value) => value is '\t' or '\n' or '\f' or '\r' or ' ';
}
