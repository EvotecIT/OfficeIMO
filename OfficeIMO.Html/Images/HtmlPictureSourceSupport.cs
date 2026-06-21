namespace OfficeIMO.Html;

/// <summary>
/// Shared conversion support rules for picture source content types.
/// </summary>
internal static class HtmlPictureSourceSupport {
    /// <summary>
    /// Returns whether a picture source type can be consumed by the shared OfficeIMO conversion adapters.
    /// </summary>
    internal static bool IsSupportedConversionContentType(string? type) {
        if (string.IsNullOrWhiteSpace(type)) {
            return true;
        }

        string normalized = type!.Trim();
        int parameterStart = normalized.IndexOf(';');
        if (parameterStart >= 0) {
            normalized = normalized.Substring(0, parameterStart).Trim();
        }

        switch (normalized.ToLowerInvariant()) {
            case "image/apng":
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
}
