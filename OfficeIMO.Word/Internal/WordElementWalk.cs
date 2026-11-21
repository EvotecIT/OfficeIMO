using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Pre-order walk of an element's descendants without the iterator and allocation of <c>Descendants()</c>.</summary>
internal static class WordElementWalk {
    /// <summary>Paragraph and run formatting and text, which hold no block content, content controls or tables.</summary>
    /// <remarks>Only decides whether a walk descends, so exact types suffice: a type test against classes that are not
    /// sealed costs more than the rest of the step.</remarks>
    internal static bool IsFormattingOrText(OpenXmlElement element) {
        Type type = element.GetType();
        return type == typeof(ParagraphProperties) || type == typeof(RunProperties) || type == typeof(Text);
    }

    /// <summary>
    /// Whether <paramref name="element"/> is one of the most common body elements, which are none of the tables,
    /// cells and cell-merge markers that walks look for. It lets a walk skip type tests on most elements.
    /// </summary>
    internal static bool IsCommonContent(OpenXmlElement element) {
        Type type = element.GetType();
        return type == typeof(Paragraph) || type == typeof(Run) || type == typeof(Text) ||
            type == typeof(ParagraphProperties) || type == typeof(RunProperties) || type == typeof(TableRow) ||
            type == typeof(TableCellProperties) || type == typeof(TableCellWidth);
    }

    /// <summary>
    /// Returns the first child of type <typeparamref name="T"/> by following sibling links, without the iterator of
    /// <c>GetFirstChild</c> or the schema lookup of the typed child properties.
    /// </summary>
    internal static T? FirstChild<T>(OpenXmlElement parent) where T : OpenXmlElement {
        for (OpenXmlElement? child = parent.FirstChild; child != null; child = child.NextSibling()) {
            if (child is T match) return match;
        }
        return null;
    }

    /// <summary>Returns the element after <paramref name="element"/> in document order inside <paramref name="root"/>.</summary>
    /// <param name="element">Current descendant of <paramref name="root"/>.</param>
    /// <param name="root">Element whose descendants are walked.</param>
    /// <param name="descend">Whether to visit the children of <paramref name="element"/>.</param>
    internal static OpenXmlElement? Next(OpenXmlElement element, OpenXmlElement root, bool descend) {
        if (descend && element.FirstChild is { } child) return child;
        for (OpenXmlElement? current = element; current != null && !ReferenceEquals(current, root); current = current.Parent) {
            if (current.NextSibling() is { } sibling) return sibling;
        }
        return null;
    }
}
