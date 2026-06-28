namespace OfficeIMO.Markdown;

internal static class HtmlTextEncoder {
    internal static string Encode(string? text) {
        if (string.IsNullOrEmpty(text)) {
            return string.Empty;
        }

        var value = text!;
        var builder = new System.Text.StringBuilder(value.Length);
        for (var i = 0; i < value.Length; i++) {
            switch (value[i]) {
                case '&':
                    builder.Append("&amp;");
                    break;
                case '<':
                    builder.Append("&lt;");
                    break;
                case '>':
                    builder.Append("&gt;");
                    break;
                case '"':
                    builder.Append("&quot;");
                    break;
                case '\'':
                    builder.Append("&#39;");
                    break;
                default:
                    builder.Append(value[i]);
                    break;
            }
        }

        return builder.ToString();
    }
}
