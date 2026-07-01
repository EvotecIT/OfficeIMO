namespace OfficeIMO.Markdown;

internal static class MarkdownAttributeBlockRenderer {
    internal static string RenderTrailing(MarkdownAttributeSet? attributes) {
        return Render(attributes, includeLeadingSpace: true);
    }

    internal static string RenderInlineTrailing(MarkdownAttributeSet? attributes) {
        return Render(attributes, includeLeadingSpace: false);
    }

    private static string Render(MarkdownAttributeSet? attributes, bool includeLeadingSpace) {
        if (attributes == null || attributes.IsEmpty) {
            return string.Empty;
        }

        var builder = new StringBuilder();
        if (includeLeadingSpace) {
            builder.Append(' ');
        }

        builder.Append('{');

        var hasToken = false;
        if (!string.IsNullOrWhiteSpace(attributes.ElementId)) {
            AppendSeparator(builder, ref hasToken);
            builder.Append('#').Append(attributes.ElementId!.Trim());
        }

        for (int i = 0; i < attributes.Classes.Count; i++) {
            if (string.IsNullOrWhiteSpace(attributes.Classes[i])) {
                continue;
            }

            AppendSeparator(builder, ref hasToken);
            builder.Append('.').Append(attributes.Classes[i].Trim());
        }

        foreach (var attribute in attributes.Attributes.OrderBy(static pair => pair.Key, StringComparer.OrdinalIgnoreCase)) {
            if (string.IsNullOrWhiteSpace(attribute.Key)
                || string.Equals(attribute.Key, "id", StringComparison.OrdinalIgnoreCase)
                || string.Equals(attribute.Key, "class", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            AppendSeparator(builder, ref hasToken);
            builder.Append(attribute.Key.Trim());
            if (!string.Equals(attribute.Value, "true", StringComparison.OrdinalIgnoreCase)) {
                builder.Append('=').Append(Quote(attribute.Value));
            }
        }

        if (!hasToken) {
            return string.Empty;
        }

        builder.Append('}');
        return builder.ToString();
    }

    private static void AppendSeparator(StringBuilder builder, ref bool hasToken) {
        if (hasToken) {
            builder.Append(' ');
        }

        hasToken = true;
    }

    private static string Quote(string? value) {
        if (value == null) {
            return "\"\"";
        }

        return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }
}
