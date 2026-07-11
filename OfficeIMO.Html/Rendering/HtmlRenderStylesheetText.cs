using System.Text;

namespace OfficeIMO.Html;

internal static class HtmlRenderStylesheetText {
    internal static bool TryDecode(byte[] bytes, out string css) {
        css = string.Empty;
        try {
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) {
                css = new UTF8Encoding(false, true).GetString(bytes, 3, bytes.Length - 3);
            } else if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) {
                css = new UnicodeEncoding(false, true, true).GetString(bytes, 2, bytes.Length - 2);
            } else if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) {
                css = new UnicodeEncoding(true, true, true).GetString(bytes, 2, bytes.Length - 2);
            } else {
                css = new UTF8Encoding(false, true).GetString(bytes);
            }

            return true;
        } catch (DecoderFallbackException) {
            css = string.Empty;
            return false;
        }
    }
}
