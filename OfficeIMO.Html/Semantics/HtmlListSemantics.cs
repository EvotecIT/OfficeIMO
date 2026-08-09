using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

/// <summary>Canonical ordered, unordered, and definition-list interpretation.</summary>
internal static class HtmlListSemantics {
    internal static HtmlSemanticList BuildList(IElement list) {
        if (Is(list, "ol")) {
            bool reversed = list.HasAttribute("reversed");
            int defaultStart = reversed ? list.Children.Count(child => Is(child, "li")) : 1;
            int start = TryReadInteger(list.GetAttribute("start"), out int authoredStart) ? authoredStart : defaultStart;
            return new HtmlSemanticList(HtmlSemanticListKind.Ordered, start, reversed);
        }
        if (Is(list, "dl")) return new HtmlSemanticList(HtmlSemanticListKind.Definition, null, false);
        return new HtmlSemanticList(HtmlSemanticListKind.Unordered, null, false);
    }

    internal static IReadOnlyList<HtmlListItemProjection> BuildItems(IElement list, HtmlSemanticList semantics) {
        var result = new List<HtmlListItemProjection>();
        int current = semantics.Start ?? 1;
        int step = semantics.IsReversed ? -1 : 1;
        foreach (IElement element in list.Children) {
            if (semantics.Kind == HtmlSemanticListKind.Definition) {
                if (Is(element, "dt")) {
                    result.Add(new HtmlListItemProjection(element,
                        new HtmlSemanticListItem(HtmlSemanticListItemKind.Term, null, null)));
                } else if (Is(element, "dd")) {
                    result.Add(new HtmlListItemProjection(element,
                        new HtmlSemanticListItem(HtmlSemanticListItemKind.Description, null, null)));
                }
                continue;
            }
            if (!Is(element, "li")) continue;

            int? explicitOrdinal = null;
            int? ordinal = null;
            if (semantics.Kind == HtmlSemanticListKind.Ordered) {
                if (TryReadInteger(element.GetAttribute("value"), out int authoredOrdinal)) {
                    current = authoredOrdinal;
                    explicitOrdinal = authoredOrdinal;
                }
                ordinal = current;
                current += step;
            }
            result.Add(new HtmlListItemProjection(element,
                new HtmlSemanticListItem(HtmlSemanticListItemKind.Item, ordinal, explicitOrdinal)));
        }
        return result;
    }

    internal static string BuildText(HtmlSemanticList list, IReadOnlyList<HtmlSemanticBlock> items) {
        return string.Join("\n", items.Select(item => {
            if (item.ListItem?.Kind == HtmlSemanticListItemKind.Term) return item.Text;
            if (item.ListItem?.Kind == HtmlSemanticListItemKind.Description) return "  " + item.Text;
            if (list.Kind == HtmlSemanticListKind.Ordered) {
                return (item.ListItem?.Ordinal ?? 1).ToString(CultureInfo.InvariantCulture) + ". " + item.Text;
            }
            return "• " + item.Text;
        }));
    }

    private static bool TryReadInteger(string? text, out int value) =>
        int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

    private static bool Is(IElement element, string name) =>
        string.Equals(element.LocalName, name, StringComparison.OrdinalIgnoreCase);
}

internal sealed class HtmlListItemProjection {
    internal HtmlListItemProjection(IElement element, HtmlSemanticListItem semantics) {
        Element = element;
        Semantics = semantics;
    }

    internal IElement Element { get; }
    internal HtmlSemanticListItem Semantics { get; }
}
