namespace OfficeIMO.Markdown;

internal static class MarkdownHtmlAttributes {
    internal static string Render(
        MarkdownAttributeSet? attributes,
        HtmlOptions? options,
        string? fallbackElementId = null,
        IReadOnlyList<string>? additionalClasses = null,
        bool additionalClassesFirst = false) {
        if ((attributes == null || attributes.IsEmpty)
            && string.IsNullOrWhiteSpace(fallbackElementId)
            && (additionalClasses == null || additionalClasses.Count == 0)) {
            return string.Empty;
        }

        var effectiveOptions = options ?? HtmlRenderContext.Options;
        var builder = new StringBuilder();
        AppendId(builder, attributes?.ElementId ?? fallbackElementId, effectiveOptions);
        if (attributes != null || additionalClasses != null) {
            AppendClasses(builder, attributes?.Classes, additionalClasses, effectiveOptions, additionalClassesFirst);
        }

        if (attributes != null) {
            AppendAttributes(builder, attributes, effectiveOptions);
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
        HtmlOptions? options,
        bool additionalClassesFirst) {
        if ((classes == null || classes.Count == 0)
            && (additionalClasses == null || additionalClasses.Count == 0)) {
            return;
        }

        var normalized = new List<string>();
        if (additionalClassesFirst) {
            AppendNormalizedClasses(additionalClasses, normalized);
            AppendNormalizedClasses(classes, normalized);
        } else {
            AppendNormalizedClasses(classes, normalized);
            AppendNormalizedClasses(additionalClasses, normalized);
        }

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
                || !IsSafeAttributeName(attribute.Key, options)) {
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

    private static bool IsSafeAttributeName(string name, HtmlOptions? options) {
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

        if ((options?.RawHtmlHandling ?? RawHtmlHandling.Allow) == RawHtmlHandling.Allow) {
            return true;
        }

        if (trimmed.StartsWith("on", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("xmlns", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        switch (trimmed.ToLowerInvariant()) {
            case "action":
            case "background":
            case "data":
            case "formaction":
            case "href":
            case "imagesrcset":
            case "manifest":
            case "ping":
            case "poster":
            case "src":
            case "srcdoc":
            case "srcset":
            case "style":
            case "xlink:href":
                return false;
        }

        return true;
    }
}
