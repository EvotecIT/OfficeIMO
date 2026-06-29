namespace OfficeIMO.Markdown;

internal static class MarkdownHtmlAttributes {
    internal static string Render(
        MarkdownAttributeSet? attributes,
        HtmlOptions? options,
        string? fallbackElementId = null,
        IReadOnlyList<string>? additionalClasses = null) {
        if ((attributes == null || attributes.IsEmpty)
            && string.IsNullOrWhiteSpace(fallbackElementId)
            && (additionalClasses == null || additionalClasses.Count == 0)) {
            return string.Empty;
        }

        var builder = new StringBuilder();
        AppendId(builder, attributes?.ElementId ?? fallbackElementId, options);
        if (attributes != null || additionalClasses != null) {
            AppendClasses(builder, attributes?.Classes, additionalClasses, options);
        }

        if (attributes != null) {
            AppendAttributes(builder, attributes, options);
        }

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

    private static void AppendClasses(
        StringBuilder builder,
        IReadOnlyList<string>? classes,
        IReadOnlyList<string>? additionalClasses,
        HtmlOptions? options) {
        if ((classes == null || classes.Count == 0)
            && (additionalClasses == null || additionalClasses.Count == 0)) {
            return;
        }

        var normalized = new List<string>();
        AppendNormalizedClasses(classes, normalized);
        AppendNormalizedClasses(additionalClasses, normalized);

        if (normalized.Count == 0) {
            return;
        }

        builder.Append(" class=\"")
            .Append(HtmlTextEncoder.Encode(string.Join(" ", normalized), options))
            .Append('"');
    }

    private static void AppendNormalizedClasses(IReadOnlyList<string>? classes, List<string> normalized) {
        if (classes == null) {
            return;
        }

        for (int i = 0; i < classes.Count; i++) {
            if (!string.IsNullOrWhiteSpace(classes[i])) {
                normalized.Add(classes[i].Trim());
            }
        }
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
