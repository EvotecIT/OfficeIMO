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
        foreach (IElement element in EnumerateItemElements(list, semantics.Kind)) {
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
                current = Advance(current, step);
            }
            result.Add(new HtmlListItemProjection(element,
                new HtmlSemanticListItem(HtmlSemanticListItemKind.Item, ordinal, explicitOrdinal)));
        }
        return result;
    }

    private static IEnumerable<IElement> EnumerateItemElements(
        IElement list,
        HtmlSemanticListKind kind) {
        foreach (IElement child in list.Children) {
            if (kind == HtmlSemanticListKind.Definition && Is(child, "div")) {
                foreach (IElement groupedChild in child.Children) {
                    if (Is(groupedChild, "dt") || Is(groupedChild, "dd")) yield return groupedChild;
                }
            } else {
                yield return child;
            }
        }
    }

    internal static string BuildText(HtmlSemanticList list, IReadOnlyList<HtmlSemanticBlock> items) {
        return string.Join("\n", items.Select(item => {
            string ownText;
            if (item.ListItem?.Kind == HtmlSemanticListItemKind.Term) ownText = item.Text;
            else if (item.ListItem?.Kind == HtmlSemanticListItemKind.Description) ownText = "  " + item.Text;
            else if (list.Kind == HtmlSemanticListKind.Ordered) {
                ownText = (item.ListItem?.Ordinal ?? 1).ToString(CultureInfo.InvariantCulture) + ". " + item.Text;
            } else {
                ownText = "• " + item.Text;
            }

            string nestedText = string.Join("\n", item.Children
                .Where(child => child.Kind == HtmlSemanticBlockKind.List && child.Text.Length > 0)
                .Select(child => "  " + child.Text.Replace("\n", "\n  ")));
            return nestedText.Length == 0 ? ownText : ownText + "\n" + nestedText;
        }));
    }

    internal static bool TryResolveOrdinal(IElement item, out int ordinal) {
        ordinal = 0;
        IElement? list = item.ParentElement;
        if (list == null || !Is(list, "ol") || !Is(item, "li")) return false;

        HtmlSemanticList semantics = BuildList(list);
        HtmlListItemProjection? projection = BuildItems(list, semantics)
            .FirstOrDefault(candidate => ReferenceEquals(candidate.Element, item));
        if (projection == null || !projection.Semantics.Ordinal.HasValue) return false;

        ordinal = projection.Semantics.Ordinal.Value;
        return true;
    }

    private static int Advance(int value, int step) {
        if (step > 0) return value == int.MaxValue ? int.MaxValue : value + 1;
        return value == int.MinValue ? int.MinValue : value - 1;
    }

    private static bool TryReadInteger(string? text, out int value) {
        value = 0;
        if (string.IsNullOrEmpty(text)) return false;

        int position = 0;
        while (position < text!.Length && IsAsciiWhitespace(text[position])) position++;
        bool negative = false;
        if (position < text.Length && (text[position] == '+' || text[position] == '-')) {
            negative = text[position] == '-';
            position++;
        }

        int firstDigit = position;
        long magnitude = 0L;
        long limit = negative ? (long)int.MaxValue + 1L : int.MaxValue;
        while (position < text.Length && text[position] >= '0' && text[position] <= '9') {
            int digit = text[position] - '0';
            magnitude = magnitude > (limit - digit) / 10L ? limit : magnitude * 10L + digit;
            position++;
        }
        if (position == firstDigit) return false;

        value = negative
            ? magnitude == (long)int.MaxValue + 1L ? int.MinValue : -(int)magnitude
            : (int)magnitude;
        return true;
    }

    private static bool IsAsciiWhitespace(char value) =>
        value == '\t' || value == '\n' || value == '\f' || value == '\r' || value == ' ';

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
