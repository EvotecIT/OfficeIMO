using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Single DOM-based owner for locating semantic features used by preflight and scoring.</summary>
internal static class HtmlSemanticFeatureLocator {
    internal static IReadOnlyList<IElement> FindComments(IHtmlDocument document) =>
        document.QuerySelectorAll(".officeimo-comments li, [data-officeimo-comment]")
            .Distinct().ToList();

    internal static IReadOnlyList<IElement> FindAnnotations(IHtmlDocument document) =>
        document.QuerySelectorAll("ins, del, [data-officeimo-annotation], [data-officeimo-bookmark]")
            .Distinct().ToList();

    internal static IReadOnlyList<IElement> FindFormulas(IHtmlDocument document) {
        IReadOnlyList<IElement> inventory = document.QuerySelectorAll(".officeimo-formulas li").ToList();
        if (inventory.Count > 0) return inventory;
        return document.QuerySelectorAll("[data-officeimo-formula], [data-officeimo-value-kind='formula' i]").ToList();
    }

    internal static IReadOnlyList<IElement> FindCharts(IHtmlDocument document) {
        var result = new List<IElement>();
        IReadOnlyList<IElement> inventory = document.QuerySelectorAll(".officeimo-charts li").ToList();
        result.AddRange(inventory);
        foreach (IElement element in document.QuerySelectorAll("[data-officeimo-chart-type], [data-officeimo-chart-kind]")) {
            if (inventory.Any(parent => ReferenceEquals(parent, element) || parent.Contains(element))) continue;
            result.Add(element);
        }
        return result;
    }

    internal static IReadOnlyList<IElement> FindPagedLayout(IHtmlDocument document) {
        var result = new List<IElement>();
        foreach (IElement style in document.QuerySelectorAll("style")) {
            if (ContainsCssToken(style.TextContent, "@page")) result.Add(style);
        }
        foreach (IElement element in document.QuerySelectorAll("[style]")) {
            string css = element.GetAttribute("style") ?? string.Empty;
            if (ContainsCssProperty(css, "break-before", "page")
                || ContainsCssProperty(css, "break-after", "page")
                || ContainsCssProperty(css, "page-break-before", "always")
                || ContainsCssProperty(css, "page-break-after", "always")) {
                result.Add(element);
            }
        }
        return result;
    }

    private static bool ContainsCssToken(string css, string token) =>
        StripCssComments(css).IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool ContainsCssProperty(string css, string property, string value) {
        foreach (string declaration in StripCssComments(css).Split(';')) {
            int separator = declaration.IndexOf(':');
            if (separator <= 0) continue;
            string name = declaration.Substring(0, separator).Trim();
            string declared = declaration.Substring(separator + 1).Trim();
            if (string.Equals(name, property, StringComparison.OrdinalIgnoreCase)
                && declared.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
                    .Any(part => string.Equals(part.TrimEnd('!', ','), value, StringComparison.OrdinalIgnoreCase))) return true;
        }
        return false;
    }

    private static string StripCssComments(string css) {
        if (string.IsNullOrEmpty(css)) return string.Empty;
        var builder = new StringBuilder(css.Length);
        for (int index = 0; index < css.Length;) {
            if (index + 1 < css.Length && css[index] == '/' && css[index + 1] == '*') {
                int end = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                index = end < 0 ? css.Length : end + 2;
                continue;
            }
            builder.Append(css[index++]);
        }
        return builder.ToString();
    }
}
