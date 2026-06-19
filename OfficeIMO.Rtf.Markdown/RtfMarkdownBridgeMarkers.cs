using System;
using System.Globalization;

namespace OfficeIMO.Rtf.Markdown;

internal static class RtfMarkdownBridgeMarkers {
    private const string ListContinuationBookmarkPrefix = "_OfficeIMO_MLC_";
    private const string CodeBlockBookmarkPrefix = "_OfficeIMO_MCB_";

    internal static string CreateListContinuationBookmarkName(int ordinal) =>
        ListContinuationBookmarkPrefix + Math.Max(0, ordinal).ToString(CultureInfo.InvariantCulture);

    internal static bool IsListContinuationBookmark(RtfBookmarkMarker marker) =>
        marker.Name.StartsWith(ListContinuationBookmarkPrefix, StringComparison.Ordinal);

    internal static string CreateCodeBlockBookmarkName(int ordinal, string? language) =>
        CodeBlockBookmarkPrefix
        + Math.Max(0, ordinal).ToString(CultureInfo.InvariantCulture)
        + "_"
        + SanitizeBookmarkSegment(language);

    internal static bool TryGetCodeBlockBookmark(RtfBookmarkMarker marker, out string key, out string language) {
        key = string.Empty;
        language = string.Empty;
        if (marker == null || !marker.Name.StartsWith(CodeBlockBookmarkPrefix, StringComparison.Ordinal)) {
            return false;
        }

        string suffix = marker.Name.Substring(CodeBlockBookmarkPrefix.Length);
        int separator = suffix.IndexOf('_');
        if (separator < 0) {
            key = suffix;
            return true;
        }

        key = suffix.Substring(0, separator);
        language = suffix.Substring(separator + 1);
        return true;
    }

    private static string SanitizeBookmarkSegment(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        var builder = new System.Text.StringBuilder(Math.Min(value!.Length, 24));
        for (int i = 0; i < value.Length && builder.Length < 24; i++) {
            char ch = value[i];
            builder.Append(char.IsLetterOrDigit(ch) ? ch : '_');
        }

        return builder.ToString();
    }
}
