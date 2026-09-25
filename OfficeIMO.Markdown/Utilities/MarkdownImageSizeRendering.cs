using System.Globalization;

namespace OfficeIMO.Markdown;

/// <summary>Formats inline image size hints consistently for Markdown and HTML output.</summary>
internal static class MarkdownImageSizeRendering {
    internal static string ToMarkdown(double? width, double? height) {
        if (!width.HasValue && !height.HasValue) return string.Empty;
        string widthPart = width.HasValue ? "width=" + width.Value.ToString(CultureInfo.InvariantCulture) : string.Empty;
        string heightPart = height.HasValue ? "height=" + height.Value.ToString(CultureInfo.InvariantCulture) : string.Empty;
        return "{" + widthPart + (widthPart.Length > 0 && heightPart.Length > 0 ? " " : string.Empty) + heightPart + "}";
    }

    internal static string ToHtml(double? width, double? height) {
        string result = string.Empty;
        if (width.HasValue) result += " width=\"" + width.Value.ToString(CultureInfo.InvariantCulture) + "\"";
        if (height.HasValue) result += " height=\"" + height.Value.ToString(CultureInfo.InvariantCulture) + "\"";
        return result;
    }
}
