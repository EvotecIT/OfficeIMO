namespace OfficeIMO.Markdown;

internal static class MarkdownHtmlAttributes {
    internal static string Render(MarkdownAttributeSet? attributes, HtmlOptions? options) {
        if (attributes == null || attributes.IsEmpty) {
            return string.Empty;
        }

        var builder = new StringBuilder();
        AppendId(builder, attributes.ElementId, options);
        AppendClasses(builder, attributes.Classes, options);
        AppendAttributes(builder, attributes, options);
        return builder.ToString();
    }

    private static void AppendId(StringBuilder builder, string? elementId, HtmlOptions? options) {
        if (string.IsNullOrWhiteSpace(elementId)) {
            return;
        }

        builder.Append(" id=\"")
            .Append(HtmlTextEncoder.Encode(elementId!.Trim(), options))
            .Append('"');
    }

    private static void AppendClasses(StringBuilder builder, IReadOnlyList<string> classes, HtmlOptions? options) {
        if (classes.Count == 0) {
            return;
        }

        var normalized = new List<string>();
        for (int i = 0; i < classes.Count; i++) {
            if (string.IsNullOrWhiteSpace(classes[i])) {
                continue;
            }

            normalized.Add(classes[i].Trim());
        }

        if (normalized.Count == 0) {
            return;
        }

        builder.Append(" class=\"")
            .Append(HtmlTextEncoder.Encode(string.Join(" ", normalized), options))
            .Append('"');
    }

    private static void AppendAttributes(StringBuilder builder, MarkdownAttributeSet attributes, HtmlOptions? options) {
        if (attributes.Attributes.Count == 0) {
            return;
        }

        foreach (var attribute in attributes.Attributes.OrderBy(static pair => pair.Key, StringComparer.OrdinalIgnoreCase)) {
            if (string.IsNullOrWhiteSpace(attribute.Key)
                || string.Equals(attribute.Key, "id", StringComparison.OrdinalIgnoreCase)
                || string.Equals(attribute.Key, "class", StringComparison.OrdinalIgnoreCase)
                || !IsSafeAttributeName(attribute.Key)) {
                continue;
            }

            builder.Append(' ').Append(attribute.Key.Trim());
            if (attribute.Value != null) {
                builder.Append("=\"")
                    .Append(HtmlTextEncoder.Encode(attribute.Value, options))
                    .Append('"');
            }
        }
    }

    private static bool IsSafeAttributeName(string name) {
        var trimmed = name.Trim();
        if (trimmed.Length == 0) {
            return false;
        }

        for (int i = 0; i < trimmed.Length; i++) {
            char current = trimmed[i];
            if (char.IsLetterOrDigit(current)
                || current == '-'
                || current == '_'
                || current == ':'
                || current == '.') {
                continue;
            }

            return false;
        }

        return true;
    }
}
