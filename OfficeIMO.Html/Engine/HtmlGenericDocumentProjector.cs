using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

/// <summary>
/// Owns generic HTML grouping and naming decisions shared by native document adapters. It does
/// not know about Excel, PowerPoint, or OneNote artifact APIs.
/// </summary>
internal static class HtmlGenericDocumentProjector {
    private static readonly HashSet<string> IgnoredElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "script", "style", "noscript", "template", "nav"
    };

    internal static IReadOnlyList<HtmlGenericSectionProjection> CreateSections(IHtmlDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        IElement? body = document.Body ?? document.DocumentElement;
        if (body == null) return Array.Empty<HtmlGenericSectionProjection>();

        var result = new List<HtmlGenericSectionProjection>();
        IReadOnlyList<IElement> bodyBlocks = GetChildBlocks(body);
        if (bodyBlocks.Any(IsExplicitGroupingElement)) {
            AppendGroupedSections(document, body, result);
            EnsureAtLeastOneSection(document, result);
            return result;
        }

        AppendImplicitSections(document, bodyBlocks, result);
        EnsureAtLeastOneSection(document, result);
        return result;
    }

    private static void EnsureAtLeastOneSection(IHtmlDocument document, List<HtmlGenericSectionProjection> result) {
        if (result.Count > 0) return;
        string title = Normalize(document.Title);
        result.Add(new HtmlGenericSectionProjection(title.Length > 0 ? title : "Imported 1", Array.Empty<IElement>()));
    }

    private static void AppendGroupedSections(
        IHtmlDocument document,
        IElement container,
        List<HtmlGenericSectionProjection> result) {
        var pending = new List<IElement>();
        foreach (IElement child in GetChildBlocks(container)) {
            if (IgnoredElements.Contains(child.LocalName)) continue;
            if (Is(child, "section") || Is(child, "article")) {
                AppendImplicitSections(document, pending, result);
                pending.Clear();
                result.Add(new HtmlGenericSectionProjection(
                    GetSectionTitle(document, child, result.Count + 1),
                    GetChildBlocks(child)));
                continue;
            }

            if (Is(child, "main")) {
                AppendImplicitSections(document, pending, result);
                pending.Clear();
                AppendGroupedSections(document, child, result);
                continue;
            }

            pending.Add(child);
        }

        AppendImplicitSections(document, pending, result);
    }

    private static void AppendImplicitSections(
        IHtmlDocument document,
        IEnumerable<IElement> source,
        List<HtmlGenericSectionProjection> result) {
        var blocks = new List<IElement>();
        string title = Normalize(document.Title);
        foreach (IElement child in source) {
            if (IsPrimaryHeading(child) && blocks.Count > 0) {
                result.Add(new HtmlGenericSectionProjection(
                    title.Length > 0 ? title : "Imported " + (result.Count + 1).ToString(CultureInfo.InvariantCulture),
                    blocks.ToArray()));
                blocks.Clear();
                title = Normalize(child.TextContent);
                continue;
            }

            if (IsPrimaryHeading(child) && title.Length == 0) {
                title = Normalize(child.TextContent);
                continue;
            }
            blocks.Add(child);
        }

        if (blocks.Count > 0) {
            result.Add(new HtmlGenericSectionProjection(
                title.Length > 0 ? title : "Imported " + (result.Count + 1).ToString(CultureInfo.InvariantCulture),
                blocks.ToArray()));
        }
    }

    internal static IReadOnlyList<IElement> SelectRootTables(IHtmlDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return document.QuerySelectorAll("table")
            .Where(table => !HasTableAncestor(table))
            .ToList();
    }

    internal static string GetTableTitle(IHtmlDocument document, IElement table, int index) {
        string title = Normalize(table.QuerySelector(":scope > caption")?.TextContent);
        if (title.Length == 0) title = Normalize(table.GetAttribute("aria-label"));
        if (title.Length == 0) title = Normalize(table.Id);
        for (IElement? sibling = table.PreviousElementSibling; title.Length == 0 && sibling != null; sibling = sibling.PreviousElementSibling) {
            if (IsHeading(sibling)) title = Normalize(sibling.TextContent);
            if (Is(sibling, "table")) break;
        }
        if (title.Length == 0) {
            for (IElement? parent = table.ParentElement; parent != null; parent = parent.ParentElement) {
                IElement? heading = parent.Children.FirstOrDefault(IsHeading);
                if (heading != null) {
                    title = Normalize(heading.TextContent);
                    break;
                }
            }
        }
        if (title.Length == 0) title = Normalize(document.Title);
        return title.Length > 0 ? title : "Table " + index.ToString(CultureInfo.InvariantCulture);
    }

    internal static string GetBlockText(IElement element) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        if (Is(element, "ul") || Is(element, "ol")) {
            bool ordered = Is(element, "ol");
            int index = 1;
            return string.Join("\n", element.Children.Where(item => Is(item, "li")).Select(item =>
                (ordered ? (index++).ToString(CultureInfo.InvariantCulture) + ". " : "• ") + Normalize(item.TextContent)));
        }
        if (Is(element, "dl")) {
            return string.Join("\n", element.Children
                .Where(item => Is(item, "dt") || Is(item, "dd"))
                .Select(item => (Is(item, "dt") ? string.Empty : "  ") + Normalize(item.TextContent)));
        }
        if (Is(element, "pre")) return (element.TextContent ?? string.Empty).Trim();
        return Normalize(element.TextContent);
    }

    internal static IEnumerable<IElement> EnumerateBlocks(HtmlGenericSectionProjection section) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        foreach (IElement block in section.Blocks) {
            foreach (IElement candidate in EnumerateBlocks(block)) yield return candidate;
        }
    }

    internal static bool IsHeading(IElement element) =>
        element.LocalName.Length == 2
        && element.LocalName[0] == 'h'
        && element.LocalName[1] >= '1'
        && element.LocalName[1] <= '6';

    internal static bool IsTextBlock(IElement element) =>
        IsHeading(element) || Is(element, "p") || Is(element, "blockquote") || Is(element, "pre")
        || Is(element, "ul") || Is(element, "ol") || Is(element, "dl") || Is(element, "address");

    internal static bool IsTable(IElement element) => Is(element, "table");
    internal static bool IsImage(IElement element) => Is(element, "img");

    internal static bool IsMedia(IElement element) =>
        Is(element, "video") || Is(element, "audio") || Is(element, "object") || Is(element, "embed");

    internal static bool IsForm(IElement element) =>
        Is(element, "form") || Is(element, "input") || Is(element, "select")
        || Is(element, "textarea") || Is(element, "button");

    internal static bool IsNote(IElement element) =>
        HtmlAccessibilitySemantics.HasRole(element, "note")
        || HtmlAccessibilitySemantics.HasRole(element, "doc-footnote")
        || HtmlAccessibilitySemantics.HasRole(element, "doc-endnote");

    private static IEnumerable<IElement> EnumerateBlocks(IElement element) {
        if (IgnoredElements.Contains(element.LocalName)) yield break;
        if (IsTextBlock(element) || IsTable(element) || IsImage(element)
            || IsMedia(element) || IsForm(element) || IsNote(element)) {
            yield return element;
            yield break;
        }
        foreach (IElement child in GetChildBlocks(element)) {
            foreach (IElement candidate in EnumerateBlocks(child)) yield return candidate;
        }
    }

    private static IReadOnlyList<IElement> GetChildBlocks(IElement container) {
        var result = new List<IElement>();
        var text = new System.Text.StringBuilder();
        foreach (INode node in container.ChildNodes) {
            if (node is IText textNode) {
                text.Append(textNode.Data);
                continue;
            }

            if (node is not IElement element || IgnoredElements.Contains(element.LocalName)) continue;
            FlushTextBlock(container, text, result);
            result.Add(element);
        }
        FlushTextBlock(container, text, result);
        return result;
    }

    private static void FlushTextBlock(IElement container, System.Text.StringBuilder text, List<IElement> result) {
        string value = Normalize(text.ToString());
        text.Clear();
        if (value.Length == 0) return;
        IElement paragraph = container.Owner!.CreateElement("p");
        paragraph.TextContent = value;
        result.Add(paragraph);
    }

    private static string GetSectionTitle(IHtmlDocument document, IElement section, int index) {
        string title = Normalize(section.GetAttribute("aria-label"));
        if (title.Length == 0) title = Normalize(section.Children.FirstOrDefault(IsHeading)?.TextContent);
        if (title.Length == 0) title = Normalize(section.Id);
        if (title.Length == 0) title = Normalize(document.Title);
        return title.Length > 0 ? title : "Imported " + index.ToString(CultureInfo.InvariantCulture);
    }

    private static bool IsExplicitGroupingElement(IElement element) =>
        Is(element, "section") || Is(element, "article") || Is(element, "main");

    private static bool HasTableAncestor(IElement element) {
        for (IElement? parent = element.ParentElement; parent != null; parent = parent.ParentElement) {
            if (Is(parent, "table")) return true;
        }
        return false;
    }

    private static bool IsPrimaryHeading(IElement element) => Is(element, "h1") || Is(element, "h2");
    private static bool Is(IElement element, string localName) => string.Equals(element.LocalName, localName, StringComparison.OrdinalIgnoreCase);
    private static string Normalize(string? value) => string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}

internal sealed class HtmlGenericSectionProjection {
    internal HtmlGenericSectionProjection(string title, IReadOnlyList<IElement> blocks) {
        Title = title;
        Blocks = blocks;
    }

    internal string Title { get; }
    internal IReadOnlyList<IElement> Blocks { get; }
}
