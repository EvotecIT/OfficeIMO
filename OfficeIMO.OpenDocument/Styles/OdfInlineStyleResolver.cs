namespace OfficeIMO.OpenDocument;

/// <summary>Resolves text properties through nested ODF spans and hyperlinks.</summary>
internal static class OdfInlineStyleResolver {
    internal static T? Resolve<T>(OdfStyleRepository styles, XElement element, string partPath,
        Func<OdfStyle, T?> selector) where T : struct {
        foreach (OdfStyle style in InlineStyles(styles, element, partPath)) {
            T? value = selector(style);
            if (value.HasValue) return value;
        }
        return null;
    }

    internal static string? ResolveReference(OdfStyleRepository styles, XElement element, string partPath,
        Func<OdfStyle, string?> selector) {
        foreach (OdfStyle style in InlineStyles(styles, element, partPath)) {
            string? value = selector(style);
            if (value != null) return value;
        }
        return null;
    }

    internal static OdfColor? ResolveTextBackgroundColor(OdfStyleRepository styles, XElement element,
        string partPath) {
        TryResolveTextBackgroundColor(styles, element, partPath, out OdfColor? color);
        return color;
    }

    internal static bool TryResolveTextBackgroundColor(OdfStyleRepository styles, XElement element,
        string partPath, out OdfColor? color) {
        foreach (OdfStyle style in InlineStyles(styles, element, partPath)) {
            if (style.TryGetTextBackgroundColor(out color)) return true;
        }
        color = null;
        return false;
    }

    private static IEnumerable<OdfStyle> InlineStyles(OdfStyleRepository styles, XElement element,
        string partPath) {
        for (XElement? current = element; current != null && IsInline(current); current = current.Parent) {
            string? name = (string?)current.Attribute(OdfNamespaces.Text + "style-name");
            if (name == null) continue;
            OdfStyle? style = styles.FindInPart(OdfStyleFamily.Text, name, partPath)
                ?? styles.Find(OdfStyleFamily.Text, name);
            if (style == null) continue;
            foreach (OdfStyle candidate in styles.Resolve(style)) yield return candidate;
        }
    }

    private static bool IsInline(XElement element) => element.Name == OdfNamespaces.Text + "span"
        || element.Name == OdfNamespaces.Text + "a";
}
