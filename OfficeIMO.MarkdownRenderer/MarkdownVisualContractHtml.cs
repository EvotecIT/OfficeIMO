using System.Text;

namespace OfficeIMO.MarkdownRenderer;

internal static class MarkdownVisualContractHtml {
    internal const string ContractVersion = "v1";
    internal const string ConfigFormatJson = "json";
    internal const string ConfigEncodingBase64Utf8 = "base64-utf8";

    internal static MarkdownVisualPayloadData CreatePayload(string? rawContent) {
        var raw = rawContent ?? string.Empty;
        return new MarkdownVisualPayloadData(
            raw,
            Convert.ToBase64String(Encoding.UTF8.GetBytes(raw)),
            MarkdownRenderer.ComputeShortHash(raw));
    }

    internal static string BuildElementHtml(
        string elementName,
        string cssClass,
        string visualKind,
        string language,
        MarkdownVisualPayloadData payload,
        params KeyValuePair<string, string?>[] extraAttributes) {
        var sb = new StringBuilder();
        sb.Append('<')
          .Append(elementName)
          .Append(" class=\"")
          .Append(HtmlEncode(cssClass))
          .Append('"');
        AppendCommonAttributes(sb, visualKind, language, payload);
        AppendAttributes(sb, extraAttributes);
        sb.Append("></")
          .Append(elementName)
          .Append('>');
        return sb.ToString();
    }

    internal static void AppendCommonAttributes(
        StringBuilder sb,
        string visualKind,
        string language,
        MarkdownVisualPayloadData payload) {
        AppendAttribute(sb, "data-omd-visual-contract", ContractVersion);
        AppendAttribute(sb, "data-omd-visual-kind", visualKind);
        AppendAttribute(sb, "data-omd-fence-language", language);
        AppendAttribute(sb, "data-omd-visual-hash", payload.Hash);
        AppendAttribute(sb, "data-omd-config-format", ConfigFormatJson);
        AppendAttribute(sb, "data-omd-config-encoding", ConfigEncodingBase64Utf8);
        AppendAttribute(sb, "data-omd-config-b64", payload.Base64);
    }

    internal static void AppendAttribute(StringBuilder sb, string name, string? value) {
        if (string.IsNullOrWhiteSpace(name) || value == null) {
            return;
        }

        sb.Append(' ')
          .Append(name)
          .Append("=\"")
          .Append(HtmlEncode(value))
          .Append('"');
    }

    internal static void AppendAttribute(StringBuilder sb, string name, int value) {
        if (string.IsNullOrWhiteSpace(name)) {
            return;
        }

        sb.Append(' ')
          .Append(name)
          .Append("=\"")
          .Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture))
          .Append('"');
    }

    internal static void AppendAttributes(StringBuilder sb, params KeyValuePair<string, string?>[] attributes) {
        if (attributes == null || attributes.Length == 0) {
            return;
        }

        for (int i = 0; i < attributes.Length; i++) {
            var attribute = attributes[i];
            AppendAttribute(sb, attribute.Key, attribute.Value);
        }
    }

    private static string HtmlEncode(string value) {
        return System.Net.WebUtility.HtmlEncode(value ?? string.Empty);
    }
}

internal readonly struct MarkdownVisualPayloadData {
    internal MarkdownVisualPayloadData(string rawContent, string base64, string hash) {
        RawContent = rawContent ?? string.Empty;
        Base64 = base64 ?? string.Empty;
        Hash = hash ?? string.Empty;
    }

    internal string RawContent { get; }

    internal string Base64 { get; }

    internal string Hash { get; }
}
