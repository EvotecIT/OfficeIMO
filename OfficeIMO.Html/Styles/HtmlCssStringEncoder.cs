using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

/// <summary>Serializes CSS quoted strings without exposing CSS or HTML raw-text delimiters.</summary>
internal static class HtmlCssStringEncoder {
    internal static string Quote(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        var result = new StringBuilder(value.Length + 2).Append('"');
        for (int index = 0; index < value.Length; index++) {
            int scalar = char.ConvertToUtf32(value, index);
            if (char.IsHighSurrogate(value[index])) index++;
            if (scalar >= 'a' && scalar <= 'z' || scalar >= 'A' && scalar <= 'Z' ||
                scalar >= '0' && scalar <= '9' || scalar == '-' || scalar == '_') result.Append((char)scalar);
            else result.Append('\\').Append((scalar == 0 ? 0xfffd : scalar).ToString("x", CultureInfo.InvariantCulture)).Append(' ');
        }
        return result.Append('"').ToString();
    }
}
