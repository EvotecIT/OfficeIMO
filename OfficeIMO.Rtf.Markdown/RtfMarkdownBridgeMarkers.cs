using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Rtf.Markdown;

internal static class RtfMarkdownBridgeMarkers {
    private const string ListContinuationBookmarkPrefix = "_OfficeIMO_MLC_";
    private const string CodeBlockBookmarkPrefix = "_OfficeIMO_MCB_";
    private const int MaxCodeBlockInfoStringCharacters = 1_024;

    internal static string CreateListContinuationBookmarkName(int ordinal) =>
        ListContinuationBookmarkPrefix + Math.Max(0, ordinal).ToString(CultureInfo.InvariantCulture);

    internal static bool IsListContinuationBookmark(RtfBookmarkMarker marker) =>
        marker.Name.StartsWith(ListContinuationBookmarkPrefix, StringComparison.Ordinal);

    internal static string CreateCodeBlockBookmarkName(int ordinal, string? infoString) =>
        CodeBlockBookmarkPrefix
        + Math.Max(0, ordinal).ToString(CultureInfo.InvariantCulture)
        + "_"
        + EncodeBookmarkPayload(infoString);

    internal static bool TryGetCodeBlockBookmark(RtfBookmarkMarker marker, out string key, out string infoString) {
        key = string.Empty;
        infoString = string.Empty;
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
        infoString = DecodeBookmarkPayload(suffix.Substring(separator + 1));
        return true;
    }

    private static string EncodeBookmarkPayload(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        string boundedValue = value!.Length <= MaxCodeBlockInfoStringCharacters
            ? value
            : value.Substring(0, MaxCodeBlockInfoStringCharacters);
        byte[] bytes = Encoding.UTF8.GetBytes(boundedValue);
        var builder = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static string DecodeBookmarkPayload(string payload) {
        if (payload.Length == 0) {
            return string.Empty;
        }

        if ((payload.Length & 1) != 0) {
            return payload;
        }

        var bytes = new byte[payload.Length / 2];
        for (int i = 0; i < bytes.Length; i++) {
            int high = HexValue(payload[i * 2]);
            int low = HexValue(payload[(i * 2) + 1]);
            if (high < 0 || low < 0) {
                return payload;
            }

            bytes[i] = (byte)((high << 4) | low);
        }

        try {
            return Encoding.UTF8.GetString(bytes);
        } catch (DecoderFallbackException) {
            return payload;
        }
    }

    private static int HexValue(char value) {
        if (value >= '0' && value <= '9') return value - '0';
        if (value >= 'A' && value <= 'F') return value - 'A' + 10;
        if (value >= 'a' && value <= 'f') return value - 'a' + 10;
        return -1;
    }
}
